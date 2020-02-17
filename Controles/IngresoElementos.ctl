VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.UserControl IngresoElementos 
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   ScaleHeight     =   4470
   ScaleWidth      =   5850
   Begin MSMask.MaskEdBox mskDesde 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   60
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   6906
      _Version        =   393216
      Cols            =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl MMControl5 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image imgColector 
      Height          =   375
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   60
      Width           =   435
   End
   Begin VB.Image imgBuscar 
      Height          =   375
      Left            =   3900
      Stretch         =   -1  'True
      Top             =   60
      Width           =   435
   End
   Begin VB.Image imgBorrar 
      Height          =   375
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   60
      Width           =   435
   End
   Begin VB.Label lblElemento 
      Alignment       =   1  'Right Justify
      Caption         =   "Libro:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCantidad 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   2820
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblCantidadlbl 
      Alignment       =   2  'Center
      Caption         =   "Cant.:"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "IngresoElementos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_Cantidad = 0
Const m_def_Elementos = "0"
Const m_def_MostrarMensaje = 0
Const m_def_HablarMensaje = 0
Const m_def_Cod_Cliente = 0
Const m_def_Tipo_Almacenamiento = 0
Const m_def_Check_Requerimiento = 0
Const m_def_Check_Estado = 0
'Property Variables:
Dim m_Cantidad As Integer
Dim m_Elementos As Collection
Dim m_MostrarMensaje As Boolean
Dim m_HablarMensaje As Boolean
Dim m_Cod_Cliente As Integer
Rem Dim m_Tipo_Almacenamiento As tipo_almacenamiento
Dim m_Check_Requerimiento As Boolean
Dim m_Check_Estado As Boolean

'Public Sub IngresarElementos(Elementos As Variant)
'    Dim In_Elementos As New Collection
'    Dim I As Integer
'    Dim R As Integer
'    Dim c As Integer
'         If IsObject(Elementos) Then
'             Set In_Elementos = Elementos
'         Else
'              In_Elementos.Add Elementos
'         End If
'         For I = 1 To In_Elementos.Count
'             If Check_Estado Then
'                If Paradoja_Estado_Simple(In_Elementos(I), m_Cod_Cliente, Salida, m_Tipo_Almacenamiento) Then
'                    If Check_Requerimiento Then
'                       If Paradoja_Requerimientos(In_Elementos(I), m_Cod_Cliente, tipo_almacenamiento) Then
'                            CargarGrilla In_Elementos(I), Grilla, lblCantidad, True, MMControl5
'                       End If
'                    End If
'                End If
'             Else
'                CargarGrilla In_Elementos(I), Grilla, lblCantidad, True, MMControl5
'             End If
'        Next
'End Sub
'Public Property Get COD_CLIENTE() As Integer
'    COD_CLIENTE = m_Cod_Cliente
'End Property
'
'Public Property Let COD_CLIENTE(ByVal New_Cod_Cliente As Integer)
'    m_Cod_Cliente = New_Cod_Cliente
'    PropertyChanged "Cod_Cliente"
'End Property
'
''Public Property Get tipo_almacenamiento() As tipo_almacenamiento
''    tipo_almacenamiento = m_Tipo_Almacenamiento
''End Property
'
'Public Property Let tipo_almacenamiento(ByVal New_Tipo_Almacenamiento As tipo_almacenamiento)
'    m_Tipo_Almacenamiento = New_Tipo_Almacenamiento
'    PropertyChanged "Tipo_Almacenamiento"
'    Dim Titulo As String
'    Select Case New_Tipo_Almacenamiento
'    Case Caja
'        Titulo = "Cajas"
'    Case legajo
'        Titulo = "Legajos"
'    Case LIBRO
'        Titulo = "Libros"
'    End Select
'        imgBorrar.Picture = ImageList1.ListImages.Item("Borrar").Picture
'        imgBuscar.Picture = ImageList1.ListImages.Item("Buscar").Picture
'        imgColector.Picture = ImageList1.ListImages.Item("Colector").Picture
'        Grilla.TextMatrix(0, 1) = Titulo
'        Grilla.TextMatrix(0, 2) = Titulo
'        Grilla.TextMatrix(0, 3) = Titulo
'        Grilla.TextMatrix(0, 4) = Titulo
'        Grilla.TextMatrix(0, 5) = Titulo
'        Grilla.ColAlignment(1) = 4
'        Grilla.ColAlignment(2) = 4
'        Grilla.ColAlignment(3) = 4
'        Grilla.ColAlignment(4) = 4
'        Grilla.ColAlignment(5) = 4
'        lblElemento = Titulo
'End Property
'Public Property Get Check_Requerimiento() As Boolean
'    Check_Requerimiento = m_Check_Requerimiento
'End Property
'
'Public Property Let Check_Requerimiento(ByVal New_Check_Requerimiento As Boolean)
'    m_Check_Requerimiento = New_Check_Requerimiento
'    PropertyChanged "Check_Requerimiento"
'End Property
'Public Property Get Check_Estado() As Boolean
'    Check_Estado = m_Check_Estado
'End Property
'
'Public Property Let Check_Estado(ByVal New_Check_Estado As Boolean)
'    m_Check_Estado = New_Check_Estado
'    PropertyChanged "Check_Estado"
'End Property
'
'
'Private Sub imgBuscar_Click()
'    Select Case m_Tipo_Almacenamiento
'    Case Caja
'
'    Case legajo
'        frmBuscarLegajos.Show
'    Case LIBRO
'        frmBuscarLibros.Show
'    End Select
'End Sub
'
'Private Sub imgColector_Click()
'    Dim rs As New ADODB.Recordset
'    Dim Cajas As New Collection
'    Dim Sql As String
'    Sql = " SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN From LECTURACOLECTOR "
'    Sql = Sql & "Where NUMERO_LECTURA = " & InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura")
'    Sql = Sql & " ORDER BY ORDEN"
'    rs.Open Sql, CONBASA
'    Do While Not rs.EOF
'        Cajas.Add rs.Fields.Item("Caja").Value
'        If COD_CLIENTE <> rs.Fields.Item("CLIENTE").Value Then
'            MsgBox "El clinente es Incorrecto", vbCritical
'        End If
'        IngresarElementos rs.Fields.Item("Caja").Value
'        rs.MoveNext
'    Loop
'End Sub
'
'Private Sub lblCantidad_Change()
'    m_Cantidad = lblCantidad.Caption
'End Sub
'
'Private Sub mskDesde_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        IngresarElementos mskDesde.Text
'        mskDesde.Text = ""
'    End If
'End Sub
'
'Private Sub UserControl_InitProperties()
'    m_Cod_Cliente = m_def_Cod_Cliente
'    m_Tipo_Almacenamiento = m_def_Tipo_Almacenamiento
'    m_Check_Requerimiento = m_def_Check_Requerimiento
'    m_Check_Estado = m_def_Check_Estado
'    m_MostrarMensaje = m_def_MostrarMensaje
'    m_HablarMensaje = m_def_HablarMensaje
' Rem   m_Elementos = m_def_Elementos
'    m_Cantidad = m_def_Cantidad
'End Sub
'
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    m_Cod_Cliente = PropBag.ReadProperty("Cod_Cliente", m_def_Cod_Cliente)
'    m_Tipo_Almacenamiento = PropBag.ReadProperty("Tipo_Almacenamiento", m_def_Tipo_Almacenamiento)
'    m_Check_Requerimiento = PropBag.ReadProperty("Check_Requerimiento", m_def_Check_Requerimiento)
'    m_Check_Estado = PropBag.ReadProperty("Check_Estado", m_def_Check_Estado)
'    Set Grilla.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_MostrarMensaje = PropBag.ReadProperty("MostrarMensaje", m_def_MostrarMensaje)
'    m_HablarMensaje = PropBag.ReadProperty("HablarMensaje", m_def_HablarMensaje)
'   Rem  m_Elementos = PropBag.ReadProperty("Elementos", m_def_Elementos)
'    m_Cantidad = PropBag.ReadProperty("Cantidad", m_def_Cantidad)
'End Sub
'
'Private Sub UserControl_Resize()
'        Grilla.Width = UserControl.Width - 100
'        Grilla.Height = UserControl.Height - 100
'        Grilla.ColWidth(0) = 100
'        Grilla.ColWidth(1) = (Grilla.Width - Grilla.ColWidth(0)) / 5
'        Grilla.ColWidth(2) = (Grilla.Width - Grilla.ColWidth(0)) / 5
'        Grilla.ColWidth(3) = (Grilla.Width - Grilla.ColWidth(0)) / 5
'        Grilla.ColWidth(4) = (Grilla.Width - Grilla.ColWidth(0)) / 5
'        Grilla.ColWidth(5) = (Grilla.Width - Grilla.ColWidth(0)) / 5
'End Sub
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("Cod_Cliente", m_Cod_Cliente, m_def_Cod_Cliente)
'    Call PropBag.WriteProperty("Tipo_Almacenamiento", m_Tipo_Almacenamiento, m_def_Tipo_Almacenamiento)
'    Call PropBag.WriteProperty("Check_Requerimiento", m_Check_Requerimiento, m_def_Check_Requerimiento)
'    Call PropBag.WriteProperty("Check_Estado", m_Check_Estado, m_def_Check_Estado)
'    Call PropBag.WriteProperty("Font", Grilla.Font, Ambient.Font)
'    Call PropBag.WriteProperty("MostrarMensaje", m_MostrarMensaje, m_def_MostrarMensaje)
'    Call PropBag.WriteProperty("HablarMensaje", m_HablarMensaje, m_def_HablarMensaje)
'    Call PropBag.WriteProperty("Elementos", m_Elementos, m_def_Elementos)
'    Call PropBag.WriteProperty("Cantidad", m_Cantidad, m_def_Cantidad)
'End Sub
'Public Property Get Font() As Font
'    Set Font = Grilla.Font
'End Property
'Public Property Set Font(ByVal New_Font As Font)
'    Set Grilla.Font = New_Font
'    Set lblCantidad.Font = New_Font
'    Set lblElemento.Font = New_Font
'    Set lblCantidadlbl.Font = New_Font
'    Set mskDesde.Font = New_Font
'    PropertyChanged "Font"
'End Property
'Public Property Get MostrarMensaje() As Boolean
'    MostrarMensaje = m_MostrarMensaje
'End Property
'Public Property Let MostrarMensaje(ByVal New_MostrarMensaje As Boolean)
'    m_MostrarMensaje = New_MostrarMensaje
'    PropertyChanged "MostrarMensaje"
'End Property
'Public Property Get HablarMensaje() As Boolean
'    HablarMensaje = m_HablarMensaje
'End Property
'Public Property Let HablarMensaje(ByVal New_HablarMensaje As Boolean)
'    m_HablarMensaje = New_HablarMensaje
'    PropertyChanged "HablarMensaje"
'End Property
'Public Function Paradoja_Requerimientos(ELEMENTO As Variant, P_Cod_cliente As Integer, P_Tipo_Almacenamiento As tipo_almacenamiento) As Boolean
'    Dim rsRequerimientos As New ADODB.Recordset
'    Dim Ssql As String
'    Paradoja_Requerimientos = True
'    Ssql = " SELECT REQUERIMIENTO.IDREQUERIMIENTO"
'    Ssql = Ssql & vbCrLf & " FROM REQUERIMIENTOS_DETALLE, Requerimiento "
'    Ssql = Ssql & vbCrLf & " WHERE REQUERIMIENTOS_DETALLE.COD_REQUERIMIENTO = Requerimiento.IDRequerimiento"
'    Ssql = Ssql & vbCrLf & " AND (REQUERIMIENTO.ID_CLIENTE = " & P_Cod_cliente
'    Ssql = Ssql & vbCrLf & " ) AND (REQUERIMIENTO.IDESTADO < 7) "
'    Ssql = Ssql & vbCrLf & " AND (REQUERIMIENTOS_DETALLE.ELEMENTO = " & ELEMENTO & ") "
'    Ssql = Ssql & vbCrLf & " AND (REQUERIMIENTOS_DETALLE.COD_ALMACENAMIENTO =" & P_Tipo_Almacenamiento & ")"
'    rsRequerimientos.Open Ssql, CONBASA
'    Do While Not rsRequerimientos.EOF
'        Paradoja_Requerimientos = False
'        Select Case P_Tipo_Almacenamiento
'        Case Caja
'            MsgBox "La caja Numero " & ELEMENTO & " tiene un Requerimiento pendiente " & vbCrLf & " Con numero  " & rsRequerimientos!IDRequerimiento, vbCritical
'        Case legajo
'            MsgBox "El legajo Numero " & ELEMENTO & " tiene un Requerimiento pendiente " & vbCrLf & " Con numero  " & rsRequerimientos!IDRequerimiento, vbCritical
'        Case LIBRO
'            MsgBox "El libro Numero " & ELEMENTO & " tiene un Requerimiento pendiente " & vbCrLf & " Con numero  " & rsRequerimientos!IDRequerimiento, vbCritical
'        End Select
'        rsRequerimientos.MoveNext
'    Loop
'End Function
'
'Public Property Get Elementos() As Collection
'   Set Elementos = New Collection
'   Dim R, c As Integer
'   For R = 1 To Grilla.Rows - 1
'        For c = 1 To Grilla.Cols - 1
'            If Grilla.TextMatrix(R, c) <> "" Then
'                Elementos.Add Grilla.TextMatrix(R, c)
'            End If
'        Next
'   Next
' End Property
'
'Public Property Let Elementos(ByVal New_Elementos As Collection)
'    If Ambient.UserMode = False Then Err.Raise 387
'    If Ambient.UserMode Then Err.Raise 382
'   Rem  m_Elementos = New_Elementos
'    PropertyChanged "Elementos"
'End Property
'
'Public Property Get CANTIDAD() As Integer
'    CANTIDAD = m_Cantidad
'End Property
'
'Public Property Let CANTIDAD(ByVal New_Cantidad As Integer)
'    If Ambient.UserMode = False Then Err.Raise 387
'    If Ambient.UserMode Then Err.Raise 382
'    m_Cantidad = New_Cantidad
'    PropertyChanged "Cantidad"
'End Property
'
Private Sub UserControl_Initialize()
    Inicio
End Sub
