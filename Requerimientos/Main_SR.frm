VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{E435BBAF-9B5B-11D3-8204-0060089D62A8}#1.0#0"; "MiOcx1.ocx"
Begin VB.Form FormRemito1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de remitos"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5520
      TabIndex        =   30
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6780
      TabIndex        =   29
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ListBox lstPersonal 
      Height          =   2535
      ItemData        =   "Main_SR.frx":0000
      Left            =   5040
      List            =   "Main_SR.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   28
      Top             =   2100
      Width           =   2895
   End
   Begin Threed.SSPanel Nombre_Cliente 
      Height          =   585
      Left            =   1620
      TabIndex        =   11
      Top             =   540
      Width           =   5115
      _Version        =   65536
      _ExtentX        =   9022
      _ExtentY        =   1032
      _StockProps     =   15
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Alignment       =   1
   End
   Begin prjMyOcx.MyOcx_Texto Desde 
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   924
      _ExtentX        =   1640
      _ExtentY        =   529
      SoloNumero      =   -1  'True
      Text            =   "MyOcx_Texto2"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjMyOcx.MyOcx_Texto Detalle 
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   3216
      Visible         =   0   'False
      Width           =   924
      _ExtentX        =   1640
      _ExtentY        =   529
      ConSignos       =   -1  'True
      Text            =   "MyOcx_Texto1"
      ConEspacios     =   -1  'True
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      ItemData        =   "Main_SR.frx":0004
      Left            =   240
      List            =   "Main_SR.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   1404
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   2505
      Left            =   60
      TabIndex        =   5
      Top             =   2160
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   4419
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin Threed.SSPanel Separador 
      Height          =   60
      Left            =   0
      TabIndex        =   12
      Top             =   4656
      Visible         =   0   'False
      Width           =   444
      _Version        =   65536
      _ExtentX        =   783
      _ExtentY        =   106
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel PanelFecha 
      Height          =   588
      Left            =   6768
      TabIndex        =   1
      Top             =   528
      Width           =   1116
      _Version        =   65536
      _ExtentX        =   1968
      _ExtentY        =   1037
      _StockProps     =   15
      Caption         =   " &Fecha :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.29
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.TextBox Fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin Threed.SSPanel PanelIdCliente 
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   " &Id cliente :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.TextBox id_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.TextBox Observaciones 
      Enabled         =   0   'False
      Height          =   732
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   7815
   End
   Begin Threed.SSPanel MnuVolador 
      Height          =   390
      Left            =   300
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   390
      _Version        =   65536
      _ExtentX        =   677
      _ExtentY        =   677
      _StockProps     =   15
      Caption         =   "SSPanel5"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView ListaClientes 
         Height          =   135
         Left            =   60
         TabIndex        =   8
         Top             =   120
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   238
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin Threed.SSPanel PanelTitulo 
      Height          =   450
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7860
      _Version        =   65536
      _ExtentX        =   13864
      _ExtentY        =   794
      _StockProps     =   15
      Caption         =   "SISTEMA DE REMITOS"
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   19.51
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Crystal.CrystalReport cryRemito 
         Left            =   6480
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   """\\Server1basa\Sistemas\Requerimientos\remito.rpt"""
         Connect         =   """DSN = bpdc;UID = "" & UserName & "";PWD = "" & Password"
         UserName        =   "UserName "
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame BloqueOpciones 
      Enabled         =   0   'False
      Height          =   945
      Left            =   120
      TabIndex        =   10
      Top             =   1104
      Width           =   7770
      Begin VB.TextBox Fecha_Ingreso 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/M/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox Estado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Main_SR.frx":0023
         Left            =   1500
         List            =   "Main_SR.frx":002D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   1932
      End
      Begin VB.OptionButton Operacion 
         Caption         =   "Salida"
         Enabled         =   0   'False
         Height          =   252
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   924
      End
      Begin VB.Label Label2 
         Caption         =   "Numero Reque."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Cant 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5220
         TabIndex        =   25
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label lblIDRequerimiento 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5220
         TabIndex        =   23
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label lblCajaLibro 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Ingreso :"
         Height          =   204
         Left            =   96
         TabIndex        =   4
         Top             =   528
         Width           =   1404
      End
   End
   Begin Threed.SSCommand btnconsulta 
      Height          =   570
      Left            =   1320
      TabIndex        =   19
      ToolTipText     =   "Consulta de clientes"
      Top             =   540
      Width           =   630
      _Version        =   65536
      _ExtentX        =   1111
      _ExtentY        =   1005
      _StockProps     =   78
      BevelWidth      =   0
      Outline         =   0   'False
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   390
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   390
      _Version        =   65536
      _ExtentX        =   677
      _ExtentY        =   677
      _StockProps     =   15
      Caption         =   "SSPanel5"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView ListView1 
         Height          =   192
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   344
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label lblDescripcionRequerimiento 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   120
      TabIndex        =   24
      Top             =   4740
      Width           =   7815
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   48
      Top             =   4752
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_SR.frx":0042
            Key             =   "Cliente"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuRemito 
      Caption         =   "Remito"
      Visible         =   0   'False
      Begin VB.Menu MnuRemitoNuevo 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuRemitoAbrir 
         Caption         =   "Abrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuRemitoCerrar 
         Caption         =   "Cerrar"
      End
      Begin VB.Menu MnuRemitoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRemitoGuardar 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuRemitoSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRemitoImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuRemitoPrevia 
         Caption         =   "Vista previa"
      End
      Begin VB.Menu MnuRemitoCfg 
         Caption         =   "Configurar impresora"
      End
   End
   Begin VB.Menu MnuVentana 
      Caption         =   "Ventana"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu MnuVCascada 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu MnuVMosaico 
         Caption         =   "&Mosaico"
      End
      Begin VB.Menu MnuVOrg 
         Caption         =   "&Organizar iconos"
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "Salir"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FormRemito1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const GAP = 60
Const BORDER_WIDTH = 100

Const C_TIPO = 1
Const C_DESDE = 2
Const C_HASTA = 3
Const C_DETALLE = 4

Dim IgnoreModify As Boolean

Public MaxRow As Integer
Public V_Operacion As Integer
Public V_Estado As Integer
Public V_Tipo As Integer
Public NumeroRemito As Long

Dim ErrorFecha As Boolean
Dim Proximo_Nro_Remito As Long
Dim Estirar As Boolean
Dim AjustarTamaño As Boolean
Dim V_Lista As ListItem


Private Sub btnconsulta_GotFocus()
    MnuVolador.Visible = False
End Sub
Private Sub btnconsulta_Click()
  AjustarMnuVolador btnconsulta.top + btnconsulta.Height, btnconsulta.left
  LlenarLista
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

Private Sub Desde_Change()
    Grilla = Desde
    Grilla.TextMatrix(Grilla.row, 3) = Desde
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    NuevaFila Grilla
    EditKeyCode Grilla, Desde, KeyCode, Shift
End Sub

Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
    NuevaFila Grilla
    EditKeyCode Grilla, Detalle, KeyCode, Shift
End Sub

Private Sub Estado_Click()
V_Estado = Estado.ListIndex
End Sub

Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If V_Tipo = 0 Then
        
    Else
        Operacion(0).SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    TitulosGrilla
    Dim rsPersonal As OraDynaset
    Set rsPersonal = OraDatabase.CreateDynaset("Select * from Personal WHERE NAVES=1   ", ORADYN_READONLY)
    Do While Not rsPersonal.EOF
        lstPersonal.AddItem CStr(rsPersonal!IDPERSONAL) & " - " & CStr(rsPersonal!Nombre)
        rsPersonal.MoveNext
    Loop
    Screen.MousePointer = 11
    CargarRemito
    cryRemito.Connect = "DSN = bpdc;UID = " & UserName & ";PWD = " & Password
    cryRemito.ReportFileName = "\\Server1basa\Sistemas\Requerimientos\remito.rpt"
    Screen.MousePointer = 0
End Sub

Private Sub LlenarLista()
Dim clientes As OraDynaset
Dim Nombre_Cliente As String

Set clientes = OraDatabase.CreateDynaset("Select * from clientes where id_cliente<>0", ORADYN_READONLY)
    ListaClientes.SmallIcons = ImageList1
   'ListaClientes.View = lvwReport
    ListaClientes.ListItems.Clear
    Do While Not clientes.EOF
        Nombre_Cliente = clientes("Razon_Social")
        Set V_Lista = ListaClientes.ListItems.Add(, , Nombre_Cliente)
            V_Lista.SmallIcon = "Cliente"
            V_Lista.Tag = Val(clientes("id_cliente"))
        clientes.MoveNext
    Loop
    ListaClientes.ListItems.Item(1).Selected = True
End Sub
Private Sub AjustarMnuVolador(top As Long, left As Long)
    MnuVolador.top = top
    MnuVolador.left = left
    MnuVolador.Width = 2844
    MnuVolador.Height = 3084
    MnuVolador.Visible = True
    ListaClientes.top = 24 ' -230
    ListaClientes.left = 24
    ListaClientes.Height = 3025
    ListaClientes.Width = 2796
    ListaClientes.SetFocus
End Sub

Private Sub InsertaSeleccion()
    id_cliente = ListaClientes.SelectedItem.Tag
    Nombre_Cliente.Caption = ListaClientes.SelectedItem.Text
End Sub


Private Sub Grilla_DblClick()
    Select Case Grilla.Col
        Case 1
            FlexGridEdit Grilla, Tipo, vbKeySpace
        Case 2, 3
             FlexGridEdit Grilla, Desde, vbKeySpace
        Case 4
             FlexGridEdit Grilla, Detalle, vbKeySpace
    End Select
End Sub

Private Sub Grilla_GotFocus()
    Select Case Grilla.Col
    Case 1
      If Tipo.Visible = False Then Exit Sub
      Grilla = Tipo.Text
      Tipo.Visible = False
    Case 2, 3
        If Desde.Visible = False Then Exit Sub
        Grilla = Desde.Text
        If Grilla.Col = 2 Then
            Grilla.TextMatrix(Grilla.row, 3) = Desde.Text
        End If
        Desde.Visible = False
    Case 4
      If Detalle.Visible = False Then Exit Sub
      Grilla = Detalle.Text
      Detalle.Visible = False
    End Select
    Cant.Caption = Contar
    
End Sub

Private Sub Grilla_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim i As Integer
   If KeyCode = vbKeyDelete Then
      If MsgBox("uSTED ", vbYesNo) = vbYes Then
         Grilla.RemoveItem Grilla.RowSel
      End If
   End If
End Sub

Private Sub Grilla_KeyPress(KeyAscii As Integer)
    Select Case Grilla.Col
        Case 1
            FlexGridEdit Grilla, Tipo, KeyAscii
        Case 2, 3
             FlexGridEdit Grilla, Desde, KeyAscii
        Case 4
             FlexGridEdit Grilla, Detalle, KeyAscii
    End Select
End Sub

Private Sub Grilla_LeaveCell()
    Select Case Grilla.Col
    Case 1
      If Tipo.Visible = False Then Exit Sub
      Grilla = Tipo.Text
      Tipo.Visible = False
    Case 2, 3
        If Desde.Visible = False Then Exit Sub
        Grilla = Desde.Text
        If Grilla.Col = 2 Then
            Grilla.TextMatrix(Grilla.row, 3) = Desde.Text
        End If
        Desde.Visible = False
    Case 4
      If Detalle.Visible = False Then Exit Sub
      Grilla = Detalle.Text
      Detalle.Visible = False
    End Select
    Cant.Caption = Contar
End Sub

Private Sub Grilla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 0
End Sub

Function Desde_Hasta(Index As Integer) As Boolean
Dim opc As Integer
Dim Mensaje As String

Dim Desde As Double
Dim Hasta As Double

    Desde = IIf(Grilla.TextMatrix(Index, C_DESDE) = "", 0, Grilla.TextMatrix(Index, C_DESDE))
    Hasta = IIf(Grilla.TextMatrix(Index, C_HASTA) = "", 0, Grilla.TextMatrix(Index, C_HASTA))
    
    opc = Ctrl_Desde_Hasta(CDbl(Desde), CDbl(Hasta))
    
    Desde_Hasta = True
    Select Case opc
    Case 0
        Desde_Hasta = False
    Case 1 ' Desde mayor que hasta
        Mensaje = "Desde es mayor que hasta"
        ' PresentMessage de(Index), Mensaje, vbOKOnly, "Valor requerido"
        Grilla.row = Index
        MsgBox Mensaje, vbExclamation
    End Select
End Function

Private Function Ctrl_Desde_Hasta(L_Desde As Double, L_Hasta As Double) _
As Integer

' 1er Control
    If L_Desde > L_Hasta Then Ctrl_Desde_Hasta = 1: Exit Function ' desde mayor que hasta

End Function

Private Sub id_cliente_BotonClick()
    AjustarMnuVolador PanelIdCliente.top + PanelIdCliente.Height, PanelIdCliente.left
    LlenarLista
End Sub

Private Sub id_cliente_Change()
    BloqueoControles Not IsNull(id_cliente)
End Sub

Private Sub id_cliente_GotFocus()
    MnuVolador.Visible = False
End Sub

Private Sub id_cliente_KeyPress(KeyAscii As Integer)
    Dim clientes As OraDynaset
    Dim Sql As String
    Sql = "Select * from clientes where id_cliente =" + Format(Mid(id_cliente, 1, 10))
    If KeyAscii = 13 And id_cliente <> "" Then
        Set clientes = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
        If Not clientes.EOF Then
            Nombre_Cliente.Caption = clientes("Razon_Social")
        End If
    End If
End Sub



Private Sub lblDescripcionRequerimiento_DblClick()
Observaciones = lblDescripcionRequerimiento
End Sub

Private Sub ListaClientes_Click()
    InsertaSeleccion
End Sub

Private Sub ListaClientes_DblClick()
    InsertaSeleccion
    MnuVolador.Visible = False
    id_cliente.SetFocus
End Sub



Private Sub MnuRemitoCerrar_Click()
    Unload Me
End Sub


Private Function NoRepetidos(Index As Integer, edt As Control) As Boolean
Dim Mensaje As String, r As Integer

Dim Tipo0 As String
Dim Desde0 As Double
Dim Hasta0 As Double

Dim Tipo1 As String
Dim Desde1 As Double
Dim Hasta1 As Double


    NoRepetidos = False
    Mensaje = "La valor está repetido, modifique este valor"
    
    Tipo1 = IIf(Grilla.TextMatrix(Index, C_TIPO) = "", 0, Grilla.TextMatrix(Index, C_TIPO))
    Desde1 = IIf(Grilla.TextMatrix(Index, C_DESDE) = "", 0, Grilla.TextMatrix(Index, C_DESDE))
    Hasta1 = IIf(Grilla.TextMatrix(Index, C_HASTA) = "", 0, Grilla.TextMatrix(Index, C_HASTA))
    
   
    For r = 1 To Grilla.Rows - 1
        Tipo0 = IIf(Grilla.TextMatrix(r, C_TIPO) = "", 0, Grilla.TextMatrix(r, C_TIPO))
        Desde0 = IIf(Grilla.TextMatrix(r, C_DESDE) = "", 0, Grilla.TextMatrix(r, C_DESDE))
        Hasta0 = IIf(Grilla.TextMatrix(r, C_HASTA) = "", 0, Grilla.TextMatrix(r, C_HASTA))

        If Desde0 <> 0 And Tipo0 + Str(Desde0) = _
        Tipo1 + Str(Desde1) And Index <> r Then
            MsgBox Mensaje, vbOKOnly, "Valor repetido"
            Grilla.TextMatrix(Index, C_DESDE) = ""
            edt.Text = ""
            Exit Function
        End If
        
        If Hasta0 <> 0 And Tipo0 + Str(Hasta0) = _
        Tipo1 + Str(Hasta1) And Index <> r Then
            MsgBox Mensaje, vbOKOnly, "Valor repetido"
            Grilla.TextMatrix(Index, C_HASTA) = ""
            edt.Text = ""
            Exit Function
        End If
        
        If Desde0 <> 0 And Tipo0 + Str(Hasta0) <> _
        Tipo0 + Str(Desde0) And Index <> r Then
            If Tipo1 = Tipo0 And Desde1 >= Desde0 And Desde1 _
            <= Hasta0 Then
                MsgBox Mensaje, vbOKOnly, "Valor repetido"
                Grilla.TextMatrix(Index, C_DESDE) = ""
                edt.Text = ""
                Exit Function
            End If
        End If
    Next
    NoRepetidos = True
End Function

Private Function Buscar_X_Cod_Bar(ctr As TextBox, _
Index As Integer, opc As Integer) As Integer
Dim Estanteria As Integer
Dim Horizontal As Integer
Dim Vertical As Integer
Dim NRO_Caja As Single
Dim Nro_Libro As Single
Dim Posicion As Integer
Dim Cod_Cliente As Integer
Dim OraContenedor As OraDynaset
Dim Sql As String

Select Case opc
    Case 0 ' Caja
        Estanteria = CInt(Mid(ctr, 1, 2))
        Horizontal = CInt(Mid(ctr, 3, 2))
        Vertical = CInt(Mid(ctr, 5, 2))
        Posicion = CInt(Mid(ctr, 7, 1))
        Cod_Cliente = CInt(Mid(ctr, 8, 4))
        NRO_Caja = CSng(Mid(ctr, 12, 6))
                
        If Not Cod_Cliente = id_cliente Then
            ctr.Text = ""
            Buscar_X_Cod_Bar = 2 'El codigo de barra no corresponde al cliente
            Exit Function
        End If
            
        Sql = "Select A.Cod_Cliente , B.Razon_Social, B.Nro_cajas, A.Estanteria, A.Horizontal,"
        Sql = Sql + "A.Vertical, A.Adelante_Atras, A.Nro_Caja, A.Estado"
        Sql = Sql + " From Contenedor A, Clientes B"
        Sql = Sql + " Where A.estanteria = " & Format(Estanteria) & " And"
        Sql = Sql + " A.horizontal = " & Format(Horizontal) & " And"
        Sql = Sql + " A.vertical = " & Format(Vertical) & " And"
        Sql = Sql + " A.adelante_atras = " & Format(Posicion) & " And"
        Sql = Sql + " B.Id_Cliente = " & Format(id_cliente)
        
        Screen.MousePointer = 11
        Set OraContenedor = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
        Screen.MousePointer = 0
        If Not OraContenedor.EOF Then
            ctr.Tag = OraContenedor("nro_caja")
            If V_Tipo = 0 Then
                ' guardia y costodia
                If OraContenedor("Estado") = 4 Then
                   Buscar_X_Cod_Bar = 0
                   ctr = NRO_Caja
                Else
                   Buscar_X_Cod_Bar = 3 ' La caja no esta reservada
                End If
            Else
                ' Consulta
                If V_Operacion = 0 Then  ' Entrada
                   If OraContenedor("Estado") = 3 Then
                        Buscar_X_Cod_Bar = 0
                        ctr = NRO_Caja
                   Else
                        Buscar_X_Cod_Bar = 4 ' la caja no esta en consulta
                   End If
                Else                     ' Salida
                   If OraContenedor("Estado") = 2 Then
                        Buscar_X_Cod_Bar = 0
                        ctr = NRO_Caja
                   Else
                        Buscar_X_Cod_Bar = 5 ' La caja no esta en planta para consulta
                   End If
                End If
            End If
        Else
            Buscar_X_Cod_Bar = 1 ' La caja no existe
        End If
    Case 1 ' Libro 000050000500004
    
        Nro_Libro = CSng(Mid(ctr, 6, 5))
        Cod_Cliente = CInt(Mid(ctr, 11, 5))
        
        If Not Cod_Cliente = id_cliente Then
            'MsgBox "El codigo de barra no corresponde al cliente", vbCritical, "Atención"
            ctr = ""
            Buscar_X_Cod_Bar = 2
            Exit Function
        End If
        
        Sql = "Select * From Libros"
        Sql = Sql + " Where Cod_Cliente = " + Format(id_cliente)
        Sql = Sql + " and Nro_Libro_Interno = " + Format(Nro_Libro)
        
        Screen.MousePointer = 11
        Set OraContenedor = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
        Screen.MousePointer = 0
        If Not OraContenedor.EOF Then
            ctr.Tag = OraContenedor("nro_Libro_Interno")
            If V_Tipo = 0 Then
                ' guardia y costodia
                If OraContenedor("Estado") = 4 Then
                   Buscar_X_Cod_Bar = 0
                   ctr = Nro_Libro
                Else
                    Buscar_X_Cod_Bar = 3
                End If
            Else
                ' Consulta
                If V_Operacion = 0 Then  ' Entrada
                   If OraContenedor("Estado") = 3 Then
                        Buscar_X_Cod_Bar = 0
                        ctr = Nro_Libro
                   Else
                        Buscar_X_Cod_Bar = 4
                   End If
                Else                     ' Salida
                   If OraContenedor("Estado") = 2 Then
                        Buscar_X_Cod_Bar = 0
                        ctr = Nro_Libro
                   Else
                        Buscar_X_Cod_Bar = 5
                   End If
                End If
            End If
        Else
            Buscar_X_Cod_Bar = 1
        End If
    End Select
End Function

Private Function Buscar_X_NCaja(ctr As Control, _
opc As Integer) As Integer
Dim OraContenedor As OraDynaset
Dim OraRango As OraDynaset
Dim Sql As String
Dim Msg As String
Dim Desde As Double
Dim Hasta As Double
Dim r As Single
On Error GoTo OraError

    Select Case Mid(lblCajaLibro, 1, 1)
    Case 0 ' Caja
        
        
        Desde = Grilla.TextMatrix(Grilla.row, 2)
        Hasta = Grilla.TextMatrix(Grilla.row, 3)
        Buscar_X_NCaja = 1
        For r = Desde To Hasta
            Sql = "Select * From Contenedor "
            Sql = Sql + " Where Nro_Caja = " & r & " And " + _
            " Cod_Cliente = " & Format(id_cliente)
            
            Screen.MousePointer = 11
            Set OraContenedor = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
            Screen.MousePointer = 0
            
            If Not OraContenedor.EOF Then
                If V_Tipo = 0 Then
                    ' guardia y costodia
                    If Val(OraContenedor("Estado")) = 4 Then
                        Buscar_X_NCaja = 0
                    Else
                        Buscar_X_NCaja = IIf(OraContenedor.RecordCount = 1, 3, 6)
                        If MsgBox(" La caja " & r & " Tieme problema" & vbCrLf & "Usted quiere ver los Movimientos de esta caja", vbYesNo + vbInformation) = vbYes Then
                        
                                
                                
                          End If
                        Exit Function
                    End If
                Else
                    ' Consulta
                    If V_Operacion = 0 Then  ' Entrada
                       If Val(OraContenedor("Estado")) = 3 Then
                            Buscar_X_NCaja = 0 ' si
                       Else
                            Buscar_X_NCaja = IIf(Desde = Hasta, 4, 8) ' No
                            If MsgBox("La caja " & r & " Tieme problemas " & vbCrLf & "Usted quiere ver los Movimientos de esta caja", vbYesNo) = vbYes Then
                               
                            End If
                            Exit Function
                       End If
                    Else                     ' Salida
                       If Val(OraContenedor("Estado")) = 2 Then
                            Buscar_X_NCaja = 0 ' si
                       Else
                                                                            
                            Exit Function
                       End If
                    End If
                End If
            End If
        Next
    Case 1 ' Libros
        Sql = "Select * From Libros "
        Sql = Sql + " Where Nro_Libro_Interno = " + Grilla.TextMatrix(Grilla.row, 2)
        Sql = Sql + " and Cod_Cliente = " & Format(id_cliente)
        Set OraContenedor = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
        
        If OraContenedor.EOF Then
            Buscar_X_NCaja = 1
            Exit Function
        End If
        
        Sql = "Select * From Libros "
        Sql = Sql + " Where Nro_Libro_Interno = " + Grilla.TextMatrix(Grilla.row, 3)
        Sql = Sql + " and Cod_Cliente = " & Format(id_cliente)
        Set OraContenedor = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
        
        If OraContenedor.EOF Then
            Buscar_X_NCaja = 1
            Exit Function
        End If
        
               
        Sql = "Select * From Libros "
        Sql = Sql + " Where Nro_Libro_Interno Between " + Grilla.TextMatrix(Grilla.row, 2)
        Sql = Sql + " And " + Grilla.TextMatrix(Grilla.row, 3) + " And "
        Sql = Sql + " Cod_Cliente = " & Format(id_cliente)
    
        Screen.MousePointer = 11
        Set OraContenedor = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
        Screen.MousePointer = 0
        
        Buscar_X_NCaja = 1
        Do While Not OraContenedor.EOF
            Buscar_X_NCaja = 0
            ctr.Tag = Val(OraContenedor("Nro_Libro_Interno"))
            If V_Tipo = 0 Then
                ' guardia y costodia
                If Val(OraContenedor("Estado")) = 4 Then
                    Buscar_X_NCaja = 0
                Else
                    Buscar_X_NCaja = IIf(OraContenedor.RecordCount = 1, 3, 6)
                    If MsgBox("El Libro  " & Val(OraContenedor("Nro_Libro_Interno")) & " Tieme problemas " & vbCrLf & "Usted quiere ver los Movimientos de esta caja", vbYesNo) = vbYes Then
                       
                    End If
                    Exit Do
                End If
            Else
                ' Consulta
                If V_Operacion = 0 Then  ' Entrada
                   If Val(OraContenedor("Estado")) = 3 Then
                        Buscar_X_NCaja = 0
                   Else
                        Buscar_X_NCaja = IIf(OraContenedor.RecordCount = 1, 4, 8)
                        If MsgBox("El libro " & Val(OraContenedor("Nro_Libro_Interno")) & " Tieme problemas " & vbCrLf & "Usted quiere ver los Movimientos de esta caja", vbYesNo) = vbYes Then
                           
                        End If
                        Exit Do
                   End If
                Else                     ' Salida
                   If Val(OraContenedor("Estado")) = 2 Then
                        Buscar_X_NCaja = 0
                   Else
                        Buscar_X_NCaja = IIf(OraContenedor.RecordCount = 1, 5, 10)
                        If MsgBox("El Libro " & Val(OraContenedor("Nro_Libro_Interno")) & " Tieme problemas " & vbCrLf & "Usted quiere ver los Movimientos de esta caja", vbYesNo) = vbYes Then
                            
                        End If
                        Exit Do
                   End If
                End If
            End If
            OraContenedor.MoveNext
        Loop
    Case -1 ' Error por no tener tipo
        Buscar_X_NCaja = 9
    End Select
    Exit Function
    
OraError:
    If Err = 440 Then
        Screen.MousePointer = 0
        Msg = "Hay un caracter NO valido"
        Buscar_X_NCaja = 1
        MsgBox "PresentMessage ctr, Msg, vbOKOnly, No valido"
    End If
    
End Function

Private Sub Nro_Rem_Prov_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Operacion(0).SetFocus
End Sub

Private Sub Observaciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 0
End Sub

Private Sub Operacion_Click(Index As Integer)
    Dim margen As Long
    margen = 300
    Estado.Visible = (Index = 1)
    V_Operacion = Index
    ValidControl
End Sub


Public Sub Guardar_Remito()
Dim Sql As String
Dim r As Integer
Dim Cant As Integer
Dim Mensaje As String
Dim oradyn As OraDynaset


Dim Desde As Double
Dim Hasta As Double
Dim Tipo As Integer
Dim Detalle As String

On Error GoTo OraError
Cant = Contar
Mensaje = "Está seguro de guardar los cambios"

If MsgBox(Mensaje, vbQuestion + vbYesNo, "Atención") = vbYes Then
    Screen.MousePointer = 11
    OraSession.BeginTrans
    Proximo_Nro_Remito = ProximoRemito
    
    Sql = "Insert into Remitos_Cuerpo (Nro_Remito, Nro_Rem_Prov, Tipo, Operacion,"
    Sql = Sql + " Estado, Fecha, Id_Cliente, Observaciones, Cantidad, "
    Sql = Sql + " Audit_Usuario, Audit_Fecha, Fecha_Ingreso,Fecha_Error)"
    Sql = Sql + " Values (" + Format(Proximo_Nro_Remito) + ", "         ' Nro Remito
    Sql = Sql + " '" + " " + "',"                              ' Nro remito prov.
    Sql = Sql + " TO_NUMBER('" + Format(V_Tipo) + "'), "                ' Tipo
    Sql = Sql + " TO_NUMBER('" + Format(V_Operacion) + "'), "           ' Operacion
    If V_Operacion = 1 Then
        Sql = Sql + " TO_NUMBER('" + Format(Estado.ListIndex) + "'), "  ' Estado
    Else
        Sql = Sql + " 0, "
    End If
    Sql = Sql + " TO_DATE('" & Fecha.Text & "','DD/MM/YY'), "         ' Fecha
    Sql = Sql + " TO_NUMBER('" + Format(id_cliente) + "'), "            ' Id Cliente
    Sql = Sql + " '" + Observaciones + "', "                            ' Observaciones
    Sql = Sql + " TO_NUMBER('" + Format(Cant) + "'), "                  ' Cantidad
    Sql = Sql + " '" + UCase(UserName$) + "', "                         ' Usuario
    Sql = Sql + " sysdate , " ' Fecha y Hora
    Sql = Sql + " TO_DATE ('" + Format(Fecha_Ingreso) + "','DD/MM/YY'),"
    Sql = Sql & IIf(ErrorFecha, 1, 0) & ")"
    'MsgBox Sql
    Debug.Print Sql
    OraDatabase.ExecuteSQL Sql
    
    '  cambio de Estado
        
        Dim IDPERSONAL As Integer
        Dim i As Integer
        Dim Bandera As Boolean
        Bandera = False
        Dim RS As OraDynaset
        Dim Filtro As String
        Dim FECHARECEPCION As Date
        Dim IDTIPOREQUERIMIENTO As Integer
        
            For i = 0 To lstPersonal.ListCount - 1
               IDPERSONAL = Mid(lstPersonal.List(i), 1, 2)
             If lstPersonal.Selected(i) Then
               If Bandera = True Then
                    CRequerimientos.CambioEstado IDPERSONAL, False, , 5
                Else
                    CRequerimientos.CambioEstado IDPERSONAL, True, , 5
                    Bandera = True
                End If
             End If
                
            Next
    
        Sql = "Update requerimiento set idremito = " & Proximo_Nro_Remito
        Sql = Sql + " where idrequerimiento =" & CRequerimientos.Item(1).NumeroRequerimiento
        OraDatabase.ExecuteSQL Sql
    
    
    For r = 1 To Grilla.Rows - 2
        Desde = Grilla.TextMatrix(r, C_DESDE)
        Hasta = Grilla.TextMatrix(r, C_HASTA)
        Tipo = IIf(Grilla.TextMatrix(r, C_TIPO) = "Caja", 0, 1)
        Detalle = ""
        
        If Desde <> 0 And Hasta <> 0 Then
            Tipo = Mid(lblCajaLibro.Caption, 1, 1)
            Sql = "Insert into Remitos_Detalle(Nro_Remito, Desde, Hasta,"
            Sql = Sql + " Tipo_Almacenado, Detalle, Audit_Usuario, Audit_Fecha)"
            Sql = Sql + " Values (" + Format(Proximo_Nro_Remito) + ","
            Sql = Sql + " TO_NUMBER('" + Format(Desde) + "'),"
            Sql = Sql + " TO_NUMBER('" + Format(Hasta) + "'),"
            Sql = Sql + " TO_NUMBER('" + Format(Tipo) + "'),"
            Sql = Sql & "'" & Detalle & "',"
            Sql = Sql + " '" + UCase(UserName$) + "', "                              ' Usuario
            Sql = Sql + " sysdate )"
            'MsgBox Sql
            OraDatabase.ExecuteSQL Sql
            
            GrabarMovHistorico Proximo_Nro_Remito, Desde, Hasta, id_cliente, Tipo, V_Tipo, V_Operacion, Fecha.Text
        End If
    Next
    
    Sql = "Insert into Movimientos(Id_cliente, Fecha, Nro_Remito,"
    Sql = Sql + " Tipo_Movim, Oper_Movim, Cantidad, Audit_Usuario,"
    Sql = Sql + " Audit_Fecha)"
    Sql = Sql + " Values (TO_NUMBER('" + Format(id_cliente) + "'),"     ' Id Cliente
    Sql = Sql + " TO_DATE('" & Fecha.Text & "','DD/MM/YY'), "          ' Fecha
    Sql = Sql + Format(Proximo_Nro_Remito) + ","                              ' nro remito
    Sql = Sql + " TO_NUMBER('" + Format(V_Tipo) + "'),"                 ' Tipo
    Sql = Sql + " TO_NUMBER('" + Format(V_Operacion) + "'), "           ' Operacion
    Sql = Sql + " TO_NUMBER('" + Format(Cant) + "'),"                   ' Cantidad
    Sql = Sql + " '" + UCase(UserName$) + "',"                          ' Usuario
    Sql = Sql + " sysdate)" ' Fecha de cargar
        
    ' MsgBox Sql
    OraDatabase.ExecuteSQL Sql
    
    Screen.MousePointer = 0
    
    
    If V_Tipo = 1 Then Consulta (V_Operacion)
    
    NumeroRemito = Proximo_Nro_Remito
     
    ActualizarHistorico (NumeroRemito), (Fecha.Text)
    OraSession.CommitTrans
    
    MsgBox "El remito fue grabado con exito", vbExclamation, "Remito"
    
    On Error GoTo ErrorPrn
    
    Unload Me
    
End If

Exit Sub
OraError:
    Screen.MousePointer = 0
    OraSession.Rollback
    frmLogOraError.Show MODAL
    Exit Sub
    
ErrorPrn:
    MsgBox ERROR
    Exit Sub
    
End Sub

Sub ActualizarHistorico(nrorem As Single, fec As Variant)
  Dim Sql As String
  Dim Histo As OraDynaset
  Dim Ultimo As OraDynaset

  Sql = "Select * from Historico_Remitos"
  Set Histo = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
  If Histo.EOF Then Exit Sub
  If Histo.RecordCount > 50 Then
    Sql = "Select Min(Nro_Remito) Minimo from Historico_Remitos"
    Set Ultimo = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
    Sql = "Delete from Historico_Remitos "
    Sql = Sql & " Where Nro_Remito = " & Val(Ultimo!Minimo)
    OraDatabase.ExecuteSQL Sql
  End If
  Sql = "Insert into Historico_Remitos"
  Sql = Sql & "(Nro_Remito,Fecha)"
  Sql = Sql & " Values(" & nrorem & ",to_date('" & fec & "','dd/mm/yy'))"
  OraDatabase.ExecuteSQL Sql
End Sub


Function ProximoRemito() As Long
  Dim Sql As String
  Dim OraMax As OraDynaset
  Sql = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  Set OraMax = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
  If IsNull(OraMax("Maximo")) Then ProximoRemito = 1: Exit Function
  ProximoRemito = Val(OraMax("Maximo")) + 1
End Function

Public Function ValidForm() As Boolean
Dim r As Integer
Dim Index As Integer
Dim opc As Integer
Dim i As Integer
Dim Bandera As Boolean
    ValidForm = False
    ErrorFecha = False
    Bandera = False
    
    If V_Tipo = 0 Then
         If BlankRequired(Fecha_Ingreso, "Debe introducir la fecha de inghreso al remito ") Then Exit Function
    End If
    
        For i = 0 To lstPersonal.ListCount - 1
            If lstPersonal.Selected(i) Then
                Bandera = True
            End If
        Next
    If Not Bandera Then
        MsgBox "Quien Tranporta"
        ValidForm = False
        Exit Function
    End If
    
    If BlankRequired(id_cliente, "Entre en codigo del cliente") Then Exit Function
    If BlankRequired(Fecha, "Debe introducir la fecha del remito ") Then Exit Function
    If Grilla.Rows = 2 Then Exit Function
        
    For r = 1 To Grilla.Rows - 2
        Grilla.row = r
        If Grilla.TextMatrix(r, C_TIPO) = "" Then
            Grilla.Col = C_TIPO
            Grilla.row = r
            MsgBox "Entre el tipo", vbExclamation, "Atención "
            Grilla.SetFocus
            Exit Function
        End If
        If Grilla.TextMatrix(r, C_DESDE) = "" Then
            Grilla.Col = C_DESDE
            Grilla.row = r
            MsgBox "Entre el codigo de la caja", vbExclamation, "Atención "
            Grilla.SetFocus
            Exit Function
        End If
        If Grilla.TextMatrix(r, C_HASTA) = "" Then
            MsgBox "Entre el codigo de la caja"
            Grilla.Col = C_HASTA
            Grilla.row = r
            MsgBox "Entre el codigo de la caja", vbExclamation, "Atención "
            Grilla.SetFocus
            Exit Function
        End If
        Index = IIf(Grilla.TextMatrix(Grilla.row, 1) = " Cajas", 0, 1)
        opc = Buscar_X_NCaja(Desde, Index)
        If opc Then
            MensajeError opc, Index
            Exit Function
        End If
    Next
    
    If ControlFecha Then
       ErrorFecha = True
       If MsgBox("¿ La fecha es correcta ?" + Chr(13) + _
       "Desea continua", vbQuestion + vbYesNo, "Atención") = vbNo Then
            Exit Function
       End If
    End If
    ValidForm = True
End Function
Public Function BlankRequired(ctl As Control, Msg As String) As Boolean
  If ctl.Text = "" Then
    PresentMessage ctl, Msg, vbOKOnly, "Valor requerido"
    BlankRequired = True
  Else
    BlankRequired = False
  End If
End Function
Public Sub PresentMessage(ctl As Control, Msg As String, buttons As Long, msg_title As String)
  Dim fg_color As Long
  Dim bg_color As Long

  Beep
    
  ' Remember the original colors.
  fg_color = ctl.ForeColor
  bg_color = ctl.BackColor
    
  ' Highlight the field.
  ctl.ForeColor = bg_color
  ctl.BackColor = fg_color
    
  ' Present the message.
  MsgBox Msg, buttons, msg_title
        
  ' Restore the colors.
  ctl.ForeColor = fg_color
  ctl.BackColor = bg_color
    
  ' Set the focus.
  ctl.SetFocus
End Sub
Sub ValidControl()
Dim r As Integer

    For r = 0 To 1
        Operacion(r).Enabled = True
    Next
    
    Estado.Enabled = Not V_Tipo = 2
    
    If V_Tipo = 2 Then Exit Sub
    
    If V_Tipo = 0 Then
        Operacion(1).Enabled = False
    End If


End Sub

Private Function NoFilasVacias(ctrl As TextBox) As Boolean
    NoFilasVacias = ctrl <> ""
End Function


Sub Guardia_Custodia()
Dim r As Integer
Dim Sql As String
Dim Desde
Dim Hasta
Dim Tipo
Dim Detalle

   For r = 1 To Grilla.Rows - 2
        Desde = Grilla.TextMatrix(r, C_DESDE)
        Hasta = Grilla.TextMatrix(r, C_HASTA)
        Tipo = IIf(Grilla.TextMatrix(r, C_TIPO) = "Caja", 0, 1)
        Detalle = Grilla.TextMatrix(r, C_DETALLE)
        
        If Desde <> "" And Hasta <> "" Then
            If Tipo = 0 Then   ' Cajas
                Sql = "Update Contenedor"
                Sql = Sql + " Set Estado = 2"
                Sql = Sql + " Where Cod_Cliente = " + Format(id_cliente) + " and "
                Sql = Sql + " Nro_Caja Between " + Format(Desde) + " and "
                Sql = Sql + Format(Hasta)
            ElseIf Tipo = 1 Then  ' Libros
                Sql = "Update Libros"
                Sql = Sql + " Set Estado = 2"
                Sql = Sql + " Where Cod_Cliente = " + Format(id_cliente) + " and "
                Sql = Sql + " Nro_Libro_Interno Between " + Format(Desde) + " and "
                Sql = Sql + Format(Hasta)
            End If
            'MsgBox Sql
            OraDatabase.ExecuteSQL Sql
        End If
    Next
End Sub

Sub Consulta(opc As Integer)
Dim r As Integer
Dim Sql As String
Dim Desde As Double
Dim Hasta As Double
Dim Tipo As Integer


    For r = 1 To Grilla.Rows - 2
        Desde = Grilla.TextMatrix(r, C_DESDE)
        Hasta = Grilla.TextMatrix(r, C_HASTA)
        Tipo = Mid(lblCajaLibro, 1, 1)
        If Desde <> 0 And Hasta <> 0 Then
            If Tipo = 0 Then
                Sql = "Update Contenedor"
                Sql = Sql + " Set Estado = " + IIf(opc = 0, "2", "3") '2 en planta / 3 en consulta
                Sql = Sql + " Where Cod_Cliente = " + Format(id_cliente) + " and "
                Sql = Sql + " Nro_Caja Between " + Format(Desde) + " and "
                Sql = Sql + Format(Hasta)
            ElseIf Tipo = 1 Then
                Sql = "Update Libros"
                Sql = Sql + " Set Estado = " + IIf(opc = 0, "2", "3") '2 en planta / 3 en consulta
                Sql = Sql + " Where Cod_Cliente = " + Format(id_cliente) + " and "
                Sql = Sql + " Nro_Libro_Interno Between " + Format(Desde) + " and "
                Sql = Sql + Format(Hasta)
            End If
            Debug.Print Sql
            OraDatabase.ExecuteSQL Sql
        End If
    Next
End Sub

Public Sub Actualiza_Cuerpo_Remito(OraCli As OraDynaset, oradyn As OraDynaset)
    id_cliente = Format(Val(OraCli!id_cliente), "00000")
    Nombre_Cliente.Caption = OraCli!RAZON_SOCIAL
    Fecha = oradyn!Fecha
    Fecha_Ingreso = IIf(IsNull(oradyn!Fecha_Ingreso), "", oradyn!Fecha_Ingreso)
   
    
    V_Estado = IIf(IsNull(oradyn!Estado), 0, oradyn!Estado)
    V_Operacion = oradyn!Operacion
    V_Tipo = oradyn!Tipo
    Observaciones = IIf(IsNull(oradyn("Observaciones")), "", oradyn("Observaciones"))
End Sub

Public Sub Actualiza_Detalle_Remito(oradyn As OraDynaset)
Dim r As Integer
Dim A_Tipo(1) As String
    A_Tipo(0) = "Cajas"
    A_Tipo(1) = "Libro"
    r = 1
    Do While Not oradyn.EOF
        Grilla.TextMatrix(r, C_TIPO) = A_Tipo(Val(oradyn("Tipo_Almacenado")))
        Grilla.TextMatrix(r, C_DESDE) = oradyn("Desde")
        Grilla.TextMatrix(r, C_HASTA) = oradyn("hasta")
        Grilla.TextMatrix(r, C_DETALLE) = IIf(IsNull(oradyn("detalle")), "", oradyn("detalle"))
        oradyn.MoveNext
        r = r + 1
        NuevaFila Grilla
        Grilla.row = r
    Loop
    Cant.Caption = Contar
    Operacion(V_Operacion).Value = True
    Estado.ListIndex = V_Estado

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y >= Grilla.top + Grilla.Height And Y <= Observaciones.top + 20 And _
    X >= Grilla.left And X <= Grilla.Width + Grilla.left Then
        Estirar = True
        Separador.Visible = True
        Screen.MousePointer = 7
    Else
        Separador.Visible = False
        Screen.MousePointer = 0
        Estirar = False
    End If
End Sub



Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
   NuevaFila Grilla
   EditKeyCode Grilla, Tipo, KeyCode, Shift
End Sub

Private Sub MensajeError(opc As Integer, Tipo As Integer)
Dim Msj As String
Dim obj As String
Dim Obj1 As String
    obj = IIf(Tipo = 0, "La Caja", "El Libro")
    Obj1 = IIf(Tipo = 0, "una caja", "un libro")
    Select Case opc
        Case 1
            Msj = obj + " NO existe."
        Case 2
            Msj = "El codigo de barra NO corresponde al cliente."
        Case 3
            Msj = obj + " NO está reservada."
        Case 4
            Msj = obj + " NO está en consulta."
        Case 5
            Msj = obj + " NO está en planta."
        Case 6
            Msj = "En el rango hay " + Obj1 + " que NO esta reservada."
        Case 8
            Msj = "En el rango hay " + Obj1 + " que NO está en consulta."
        Case 9
            Msj = "Tiene que cargar el tipo"
        Case 10
            Msj = "En el rango hay " + Obj1 + " que NO está en planta."
    End Select
    MsgBox Msj, vbCritical, "Atención"
End Sub

Private Function Contar() As Integer
Dim r  As Integer
Dim Desde As Double
Dim Hasta As Double

    For r = 1 To Grilla.Rows - 1
        Desde = IIf(Grilla.TextMatrix(r, C_DESDE) = "", 0, Grilla.TextMatrix(r, C_DESDE))
        Hasta = IIf(Grilla.TextMatrix(r, C_HASTA) = "", 0, Grilla.TextMatrix(r, C_HASTA))
        
        If Desde <> 0 Then
            If Desde = Hasta Then Contar = Contar + 1
            If Desde <> Hasta Then Contar = Contar + (Hasta - Desde) + 1
        End If
    Next
End Function



Sub TitulosGrilla()
Dim r As Integer, X As Integer

    Grilla.Rows = 2
    Grilla.Cols = 4
    
    ReDim Titulos(4)
    ReDim Ancho(4)
    
    Titulos(0) = "Item"
    Titulos(1) = "Tipo"
    Titulos(2) = "Desde"
    Titulos(3) = "Hasta"
   
       
    Ancho(0) = 500
    Ancho(1) = 1000
    Ancho(2) = 1000
    Ancho(3) = 1000
    
        
    For X = 0 To 1
        For r = 0 To Grilla.Cols - 1
            Grilla.TextMatrix(X, r) = IIf(X = 0, Titulos(r), "")
            Grilla.ColWidth(r) = Ancho(r)
        Next
    Next
    Grilla.Col = 1
    Grilla.row = 1
    Grilla.RowHeight(Grilla.row) = 300
    Grilla.TextMatrix(1, 0) = 1
End Sub

Sub FlexGridEdit(fg As MSFlexGrid, edt As Control, _
ByRef KeyAscii As Integer)
Dim Activar As Boolean
    Activar = True
    Select Case KeyAscii
        ' Edita el objeto actual
        Case 0 To 32
            If fg.Col = 1 And fg = "" Then
                Activar = False
            Else
                edt.Text = fg
                If fg.Col >= 2 Then edt.SelStart = Len(edt.Text)
                Activar = False
            End If
        Case Asc("0") To Asc("9")
            If fg.Col < 2 Then Exit Sub
            edt.Text = Chr(KeyAscii)
            edt.SelStart = 1
            Activar = False
        Case Else
            If fg.Col = 4 Then
                edt.Text = Chr(KeyAscii)
                edt.SelStart = 1
                Activar = False
            End If
    End Select
    If Activar Then Exit Sub
    KeyAscii = 0
    edt.Move fg.left + fg.CellLeft, _
    fg.top + fg.CellTop, fg.CellWidth - 8
    edt.Visible = True
    edt.SetFocus
End Sub

Sub EditKeyCode(ByRef fg As MSFlexGrid, edt As Control, _
KeyCode As Integer, Shift As Integer)
Dim Index As Integer, opc As Integer
    Select Case KeyCode
    Case vbKeyEscape
        edt.Visible = False
        fg.SetFocus
    Case vbKeyReturn
        fg.SetFocus
        If fg.Col = 2 Or fg.Col = 3 Then
            If fg.TextMatrix(fg.row, 1) = "Caja" Then
                Index = 0
            ElseIf fg.TextMatrix(fg.row, 1) = "Libro" Then
                Index = 1
            Else
                Index = -1
            End If
            
            opc = Buscar_X_NCaja(edt, Index)
            If Desde_Hasta(Grilla.row) Then
               fg.TextMatrix(fg.row, C_HASTA) = fg.TextMatrix(fg.row, C_DESDE)
               edt.Text = fg.TextMatrix(fg.row, C_HASTA)
               edt.SetFocus
               Exit Sub
            End If
            If opc Then
                MensajeError opc, Index
                edt.Text = ""
                fg.SetFocus
            End If
            If Not NoRepetidos(Grilla.row, edt) Then
                edt.SetFocus
                Exit Sub
            End If
        End If
        If fg.Col = 4 Then
           fg.Col = 1
           SendKeys "{Down}{Left}{Left}{Left}"
        Else
           SendKeys "{Right}"
        End If
    Case vbKeyUp
        DoEvents
        If fg.row > fg.Rows - 1 Then
            fg.row = fg.row - 1
        End If
        fg.SetFocus
    Case vbKeyDown
        DoEvents
        If fg.row < fg.Rows - 1 Then
            fg.row = fg.row + 1
        End If
        fg.SetFocus
    End Select
End Sub

Sub NuevaFila(fg As MSFlexGrid)
Dim row As Integer
    If fg.row = fg.Rows - 1 Then
        fg.Rows = fg.Rows + 1
        row = fg.Rows - 1
        fg.RowHeight(row) = 300
        fg.TextMatrix(row, 0) = row
    End If
End Sub

Sub BloqueoControles(var As Boolean)
    Grilla.Enabled = var
    
'    Operacion(1).Enabled = var
    Estado.Enabled = var
    Fecha.Enabled = var
    Observaciones.Enabled = var
    Fecha_Ingreso.Enabled = var
End Sub

Function ControlFecha() As Boolean
Dim FechaServer As OraDynaset
Dim vfecha As Variant
Dim desdef As Variant
Dim hastaf As Variant

Dim Sql As String
    Sql = "Select to_char(sysdate,'dd/mm/yy') fecha from Dual"
    ControlFecha = True
    Set FechaServer = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
    vfecha = CDate(FechaServer!Fecha)
    desdef = vfecha - 7
    hastaf = vfecha + 7
    
    If Fecha.Text >= desdef And Fecha.Text <= hastaf Then
        ControlFecha = False
    End If
    
End Function

Sub GrabarMovHistorico(mov_nrorem, mov_desde, mov_hasta, _
mov_cliente, mov_elem, mov_tipo, mov_oper, mov_fecha)
Dim r As Single
Dim Sql As String
Dim oradyn As OraDynaset
    
    Sql = "Select * from Mov_Cajas"
    Set oradyn = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
    
    For r = mov_desde To mov_hasta
        oradyn.AddNew
        oradyn!NRO_REMITO = mov_nrorem
        oradyn!NRO_Caja = r
        oradyn!id_cliente = mov_cliente
        oradyn!elemento = mov_elem
        oradyn!Tipo = mov_tipo
        oradyn!Operacion = mov_oper
        oradyn!Fecha_Movimiento = mov_fecha
        oradyn!Anulado = 0
        oradyn!Audit_Usuario = UserName
        oradyn!Audit_Fecha = Date
        oradyn.Update
    Next
    
End Sub

Public Sub CargarRemito()
 Dim Sql As String
    Dim rsRequerimiento As OraDynaset
    Dim rsRcajas As OraDynaset
    Operacion(1).Value = True
    V_Operacion = 1
    MnuRemitoGuardar.Enabled = True
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
    lblIDRequerimiento = IDREQUERIMIENTO
    If Not rsRequerimiento.EOF Then
       With rsRequerimiento
           
           Select Case !IDTIPOREQUERIMIENTO
           Case 1, 8
            lblCajaLibro = "0 Cajas"
            Estado.ListIndex = 0
            Tipo.ListIndex = 1
            V_Tipo = 1
            
           Case 2
            lblCajaLibro = "1 Libro"
            Estado.ListIndex = 1
            V_Tipo = 1
           Case 3
            lblCajaLibro = "0 Cajas"
            Estado.ListIndex = 1
            Tipo.ListIndex = 1
            V_Tipo = 1
           Case 4
            lblCajaLibro = "1 Libro"
            Estado.ListIndex = 1
            V_Tipo = 1
           End Select
            id_cliente = !id_cliente
            Nombre_Cliente = UCase(!RAZON_SOCIAL)
            Cant = CStr(!CANTIDAD)
            Fecha = Format(SysDateCompare, "dd/mm/yyyy")
            Fecha_Ingreso = Format(SysDateCompare, "dd/mm/yyyy")
            lblIDRequerimiento = CRequerimientos.Item(1).NumeroRequerimiento
            If Not IsNull(!DESCRIPCION) Then
                lblDescripcionRequerimiento = UCase(!DESCRIPCION)
            End If
            Sql = " SELECT "
            Sql = Sql & vbCrLf & " REQ.IDRequerimientos , REQ.CAJASLIBROS"
            Sql = Sql & vbCrLf & " From"
            Sql = Sql & vbCrLf & " REQUELIBOSCAJAS REQ"
            Sql = Sql & vbCrLf & " Where REQ.IDRequerimientos = " & CRequerimientos.Item(1).NumeroRequerimiento
           Set rsRcajas = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
           Do While Not rsRcajas.EOF
           
                If Not IsNull(rsRcajas!CAJASLIBROS) Then
                    Grilla.TextMatrix(Grilla.Rows - 1, 1) = Mid(lblCajaLibro, 2)
                    Grilla.TextMatrix(Grilla.Rows - 1, 2) = CStr(rsRcajas!CAJASLIBROS)
                    Grilla.TextMatrix(Grilla.Rows - 1, 3) = CStr(rsRcajas!CAJASLIBROS)
                    Grilla.AddItem ("")
                End If
           
                rsRcajas.MoveNext
           Loop
            
        End With
    End If
End Sub
Public Sub ImprimirRemito(NumeroRemito As Long)
    Dim Sql As String
    Dim sql1 As String
    Dim Bandera As Boolean
    Dim Responsables As String
    Dim RS  As OraDynaset
    Dim ANTERIOR As Long
    
    If NumeroRemito = 0 Or IsNull(NumeroRemito) Then
        Exit Sub
    End If
    
   On Error GoTo Err
    Sql = "    SELECT"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES,"
    Sql = Sql & vbCrLf & "    REMITOS_DETALLE.DESDE,"
    Sql = Sql & vbCrLf & "    REQUERIMIENTO.SECTOR, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.FECHARECEPCION,"
    Sql = Sql & vbCrLf & "    REMITO_TIPO.DESCRIPCION,"
    Sql = Sql & vbCrLf & "    REMITO_OPERACION.DESCRIPCION,"
    Sql = Sql & vbCrLf & "    REMITO_ESTADOS.DESCRIPCION,"
    Sql = Sql & vbCrLf & "    clientes.id_cliente , clientes.RAZON_SOCIAL, clientes.CALLE, clientes.NUMERO, clientes.LOCALIDAD"
    Sql = Sql & vbCrLf & "From"
    Sql = Sql & vbCrLf & "    BASA.REMITOS_CUERPO REMITOS_CUERPO,"
    Sql = Sql & vbCrLf & "    BASA.REMITOS_DETALLE REMITOS_DETALLE,"
    Sql = Sql & vbCrLf & "    BASA.REQUERIMIENTO REQUERIMIENTO,"
    Sql = Sql & vbCrLf & "    BASA.REMITO_TIPO REMITO_TIPO,"
    Sql = Sql & vbCrLf & "   BASA.REMITO_OPERACION REMITO_OPERACION,"
    Sql = Sql & vbCrLf & "   BASA.REMITO_ESTADOS REMITO_ESTADOS,"
    Sql = Sql & vbCrLf & "   BASA.clientes clientes"
    Sql = Sql & vbCrLf & " Where"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO AND"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO AND"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
    Sql = Sql & vbCrLf & "    REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO =" & NumeroRemito
           
                        sql1 = " SELECT"
            sql1 = sql1 & vbCrLf & "      H_ESTADO_REQUE.IDREQUERIMIENTO,"
            sql1 = sql1 & vbCrLf & "    H_ESTADO_REQUE.IDESTADO,"
            sql1 = sql1 & vbCrLf & "     H_ESTADO_REQUE.CONTADOR,"
            sql1 = sql1 & vbCrLf & "     PERSONAL.NOMBRE,PERSONAL.APELLIDO"
            sql1 = sql1 & vbCrLf & "  From"
            sql1 = sql1 & vbCrLf & "     BASA.H_ESTADO_REQUE , PERSONAL, Requerimiento"
            sql1 = sql1 & vbCrLf & "  Where"
            sql1 = sql1 & vbCrLf & "     H_ESTADO_REQUE.idPersonal = PERSONAL.idPersonal"
            sql1 = sql1 & vbCrLf & "     AND H_ESTADO_REQUE.IDESTADO = REQUERIMIENTO.IDESTADO"
            sql1 = sql1 & vbCrLf & "     AND H_ESTADO_REQUE.IDREQUERIMIENTO = REQUERIMIENTO.IDREQUERIMIENTO"
            sql1 = sql1 & vbCrLf & "    AND H_ESTADO_REQUE.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
            sql1 = sql1 & vbCrLf & " Order By"
            sql1 = sql1 & vbCrLf & "     H_ESTADO_REQUE.IDREQUERIMIENTO Asc"

    
    Set RS = OraDatabase.CreateDynaset(sql1, ORADYN_READONLY)
    
    Do While Not RS.EOF
        If CLng(RS!IDREQUERIMIENTO) = ANTERIOR Then
            Responsables = Responsables & " , " & CStr(RS!Nombre) & " " & CStr(RS!Apellido)
        Else
           If Bandera = False Then
                ANTERIOR = RS!IDREQUERIMIENTO
                Bandera = True
                Responsables = CStr(RS!Nombre) & " " & CStr(RS!Apellido)
           Else
Exit Do
           End If
        End If
        RS.MoveNext
    Loop
    
   DoEvents
    cryRemito.DiscardSavedData = True
    cryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
    cryRemito.Formulas(1) = "COPIA ='" & "ORIGINAL" & " '"
    cryRemito.SQLQuery = Sql
    cryRemito.Destination = 1
    cryRemito.Action = 1
    
    cryRemito.DiscardSavedData = True
    cryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
    cryRemito.Formulas(1) = "COPIA ='" & "DUPLICADO" & " '"
    cryRemito.SQLQuery = Sql
    cryRemito.Destination = 1
    cryRemito.Action = 1
    Exit Sub
Err:
    MsgBox "Atencion error al imprimir el remito " & vbCrLf & "Por favor intentolo desde la aplicacion de control de estados", vbInformation, "Error de Impresion"
End Sub
