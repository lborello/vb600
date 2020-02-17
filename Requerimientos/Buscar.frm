VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{40CE97D1-1C1F-47E7-B2C4-A9B643CAAFFD}#17.0#0"; "Controles.ocx"
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8595
   Begin Controles.ViewImg ViewImg1 
      Height          =   1455
      Left            =   540
      TabIndex        =   8
      Top             =   2520
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2566
   End
   Begin Controles.ctlClienteUsuario ctlClienteUsuario1 
      Height          =   975
      Left            =   540
      TabIndex        =   7
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1720
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
   End
   Begin MSMask.MaskEdBox mskFiltro 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   315
      Left            =   4320
      Picture         =   "Buscar.frx":0000
      TabIndex        =   3
      Top             =   480
      Width           =   1260
   End
   Begin MSComctlLib.ListView lstBusqueda 
      Height          =   3675
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   6482
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cboCriterioBusqueda 
      Height          =   315
      ItemData        =   "Buscar.frx":0442
      Left            =   1140
      List            =   "Buscar.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label lblBuscarPor 
      Caption         =   "Buscar por :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBuscar As ADODB.Recordset


Private Sub Command2_Click()
Dim MyName
Dim MyPath
Dim DateMax As Date
Dim MaxTime As Timer
Dim DateReq As Date
Dim Timereq As Date
MousePointer = 11

MyPath = "\\Recepcion\FAXserve\users\0__Administrador\"
MyName = Dir(MyPath & "*.dcx", vbDirectory)
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
    DateReq = Format(CDate(FileDateTime(MyPath & MyName)), "DD/MM/YYYY")
    Timereq = TimeValue(FileDateTime(MyPath & MyName))
    Debug.Print "Nombre: " & MyName & "  Dia:" & DateReq; " Hora:" & Timereq
      If DateReq = "17/05/2000" Then
         MsgBox "lll"
      End If
      
   End If
   MyName = Dir   ' Get next entry.
Loop
MousePointer = 0
End Sub


Private Sub cboCriterioBusqueda_Click()
'Por Nombre/Empresa del fax
'Por fecha del fax
'Por numero de Requerimiento
'Por fecha de requerimiento
'Por cliente
mskFiltro.Mask = ""
mskFiltro.Text = ""

Select Case cboCriterioBusqueda.ListIndex
Case 0
   Rem ctlClientes1.Visible = False
    mskFiltro.Visible = True
Case 1
    mskFiltro.Mask = "##/##/####"
   Rem  ctlClientes1.Visible = False
    mskFiltro.Visible = True
Case 2
    mskFiltro.Mask = "######"
   Rem  ctlClientes1.Visible = False
    mskFiltro.Visible = True
Case 3
    mskFiltro.Mask = "##/##/####"
  Rem   ctlClientes1.Visible = False
    mskFiltro.Visible = True
Case 4
 Rem    ctlClientes1.Visible = True
    mskFiltro.Visible = False
End Select
End Sub

Private Sub cmdBuscar_Click()

Dim SQL As String


Select Case cboCriterioBusqueda.ListIndex
Case 0
    SQL = "Select * from Fax"
    SQL = SQL & vbCrLf & " WHERE Nombre like '%" & UCase(mskFiltro.Text) & "%'order by fecha"
    Set rsBuscar = New ADODB.Recordset
    rsBuscar.Open SQL, ConBasa
    CargarList 2
Case 1
    mskFiltro.Mask = "##/##/####"
    SQL = "Select * from Fax"
    SQL = SQL & vbCrLf & " WHERE  " & FechaSolaString("Fecha") & "='" & mskFiltro.Text & "'"
    Set rsBuscar = New ADODB.Recordset
    rsBuscar.Open SQL, ConBasa
    CargarList 2
Case 2
    mskFiltro.Mask = "######"
    SQL = "SELECT REQUERIMIENTO.IDREQUERIMIENTO, CLIENTES.RAZON_SOCIAL,"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.IDFAX, REQUERIMIENTO.FECHARECEPCION,"
    SQL = SQL & vbCrLf & " Requerimiento.CANTIDAD , ESTADO.DESCRIPCION"
    SQL = SQL & vbCrLf & " From Requerimiento, CLIENTES, ESTADO"
    SQL = SQL & vbCrLf & " WHERE REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.IDESTADO = ESTADO.IDESTADO"
    If Trim(mskFiltro.FormattedText) <> "" Then
        SQL = SQL & vbCrLf & "  AND REQUERIMIENTO.IDREQUERIMIENTO =" & CLng(mskFiltro.FormattedText)
    End If
    SQL = SQL & vbCrLf & " ORDER BY IDREQUERIMIENTO DESC "
    Set rsBuscar = New ADODB.Recordset
    rsBuscar.Open SQL, ConBasa
    CargarList 1
Case 3
    mskFiltro.Mask = "##/##/####"
    SQL = " SELECT REQUERIMIENTO.IDREQUERIMIENTO, CLIENTES.RAZON_SOCIAL,"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.IDFAX, REQUERIMIENTO.FECHARECEPCION,"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.cantidad , ESTADO.DESCRIPCION"
    SQL = SQL & vbCrLf & " From REQUERIMIENTO, CLIENTES, ESTADO"
    SQL = SQL & vbCrLf & " WHERE REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.IDESTADO = ESTADO.IDESTADO AND"
    SQL = SQL & vbCrLf & FechaSolaString("REQUERIMIENTO.FECHARECEPCION") & "='" & mskFiltro.FormattedText & "'"
    Set rsBuscar = New ADODB.Recordset
    rsBuscar.Open SQL, ConBasa
    CargarList 1
Case 4
    SQL = " SELECT REQUERIMIENTO.IDREQUERIMIENTO, CLIENTES.RAZON_SOCIAL,"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.IDFAX, REQUERIMIENTO.FECHARECEPCION,"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.cantidad , ESTADO.DESCRIPCION"
    SQL = SQL & vbCrLf & " From REQUERIMIENTO, CLIENTES, ESTADO"
    SQL = SQL & vbCrLf & " WHERE REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.IDESTADO = ESTADO.IDESTADO AND"
    SQL = SQL & vbCrLf & " REQUERIMIENTO.ID_CLIENTE =" & ctlCliente.Valor
    SQL = SQL & vbCrLf & " order by REQUERIMIENTO.FECHARECEPCION "
    Set rsBuscar = New ADODB.Recordset
    rsBuscar.Open SQL, ConBasa
    CargarList 1
 End Select



End Sub

Private Sub lstBusqueda_DblClick()
   Dim i As Integer
 On Error GoTo LUIS:
   
   For i = 1 To lstBusqueda.ListItems.Count
        
        If lstBusqueda.ListItems(i).Selected Then
             
           If lstBusqueda.ColumnHeaders.Item(1).Text = "Nº Reque." Then
            ColocarDatosRequerimiento (CLng(lstBusqueda.ListItems(i).Text))
           Else
                ColocarDatos (CLng(lstBusqueda.ListItems(i).Text))
                Rem frmCargarRequerimientos.PonerImagen (lstBusqueda.ListItems(I).ListSubItems.Item(4).Text)
                Unload Me
                frmCargarRequerimientos.Show
                Exit Sub
           End If
        End If
   Next
Exit Sub
LUIS:
End Sub

Public Function ColocarDatos(IDFAX As Long)
'    Dim SQL As String
'    Dim rsBuscar As ADODB.Recordset
'    SQL = "select * from fax where idFax = " & IDFAX
'    Set rsBuscar = New ADODB.Recordset
'    rsBuscar.Open SQL, ConBasa
'frmCargarRequerimientos.LimpiarCampos
'frmCargarRequerimientos.lblID_fax = ""
'Do While Not rsBuscar.EOF
'    If Not IsNull(rsBuscar!DESCRIPCION) Then
'        frmCargarRequerimientos.txtMotivo = rsBuscar!DESCRIPCION
'    End If
'    frmCargarRequerimientos.maskFechafax = Format(CStr(rsBuscar!Fecha), "DD/MM/YYYY")
'    frmCargarRequerimientos.maskHorafax = Format(CStr(rsBuscar!Fecha), "HH:MM")
'    frmCargarRequerimientos.txtNombreEmpresa = rsBuscar!NOMBRE
'    frmCargarRequerimientos.fraInstitucional.Visible = True
'    Rem frmCargarRequerimientos.cboTipoComunicacion.ListIndex = 0
'    frmCargarRequerimientos.lblID_fax = rsBuscar!IDFAX
'    rsBuscar.MoveNext
'
'Loop

End Function

Public Sub CargarList(TIPO As Integer)
Dim itmX As ListItem
Dim Fecha As String
Dim SQL As String
If TIPO <> 1 Then
    
    lstBusqueda.ColumnHeaders.Clear
    lstBusqueda.ListItems.Clear
    lstBusqueda.ColumnHeaders.Add , , "Id_fax", 1000
    lstBusqueda.ColumnHeaders.Add , , "Fecha", 2000
    lstBusqueda.ColumnHeaders.Add , , "Nombre", 2000
    lstBusqueda.ColumnHeaders.Add , , "Motivo", 2000
    lstBusqueda.ColumnHeaders.Add , , "Paso", 20000
    lstBusqueda.View = lvwReport
    Do While Not rsBuscar.EOF
        If IsNull(rsBuscar!Fecha) Then
            Fecha = ""
        Else
         Fecha = CStr(rsBuscar!Fecha)
        End If
        Set itmX = lstBusqueda.ListItems.Add(, , CStr(rsBuscar!IDFAX))
        itmX.SubItems(1) = Fecha
        itmX.SubItems(2) = IIf(IsNull(rsBuscar!NOMBRE), "", Trim(rsBuscar!NOMBRE))
        itmX.SubItems(3) = IIf(IsNull(rsBuscar!DESCRIPCION), "", Trim(rsBuscar!DESCRIPCION))
        itmX.SubItems(4) = rsBuscar!Path
        rsBuscar.MoveNext
    Loop
Else
     
    
    
    lstBusqueda.ColumnHeaders.Clear
    lstBusqueda.ListItems.Clear
    
    lstBusqueda.ColumnHeaders.Add , , "Nº Reque.", 1000
    lstBusqueda.ColumnHeaders.Add , , "Fecha", 1100
    lstBusqueda.ColumnHeaders.Add , , "Cliente", 3500
    lstBusqueda.ColumnHeaders.Add , , "ID_fax", 750
    lstBusqueda.ColumnHeaders.Add , , "CatidadCajas", 700
    lstBusqueda.ColumnHeaders.Add , , "Estado", 1500
    lstBusqueda.View = lvwReport
    Do While Not rsBuscar.EOF
        If IsNull(rsBuscar!FECHARECEPCION) Then
            Fecha = ""
        Else
         Fecha = CStr(rsBuscar!FECHARECEPCION)
        End If
        Set itmX = lstBusqueda.ListItems.Add(, , CStr(rsBuscar!IDREQUERIMIENTO))
        itmX.SubItems(1) = Fecha
        itmX.SubItems(2) = IIf(IsNull(rsBuscar!Razon_Social), "", Trim(rsBuscar!Razon_Social))
        itmX.SubItems(3) = IIf(IsNull(rsBuscar!IDFAX), "", Trim(rsBuscar!IDFAX))
        itmX.SubItems(4) = rsBuscar!CANTIDAD
        itmX.SubItems(5) = rsBuscar!DESCRIPCION
        rsBuscar.MoveNext
    Loop




End If



End Sub

Public Sub ColocarDatosRequerimiento(IDREQUERIMIENTO As Long)

'Dim Sql As String
'    Dim rsRequerimiento As ADODB.Recordset
'    Dim rsRcajas As ADODB.Recordset
'
'    frmCargarRequerimientos.LimpiarCampos
'    frmCargarRequerimientos.lblID_fax = ""
'    Sql = "SELECT REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.SECTOR,"
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.TELEFONO, REQUERIMIENTO.ID_CLIENTE,"
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.DESCRIPCION, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.TOMO,"
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.FECHALIMITE, "
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.FECHARECEPCION, REQUERIMIENTO.IDTIPORECEPCION, FAX.PATH,"
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.CANTIDAD, REQUERIMIENTO.IDESTADO, REQUERIMIENTO.PEDIDOCLIENTE ,"
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.IDTIPOREQUERIMIENTO, FAX.Fecha"
'    Sql = Sql & vbCrLf & " From REQUERIMIENTO, FAX"
'    Sql = Sql & vbCrLf & " WHERE REQUERIMIENTO.IDFAX = FAX.IDFAX  (+) AND"
'    Sql = Sql & vbCrLf & " REQUERIMIENTO.IDREQUERIMIENTO = " & IDREQUERIMIENTO
'    Set rsRequerimiento = New ADODB.Recordset
'    rsRequerimiento.Open Sql, ConBasa
'
'    If Not rsRequerimiento.EOF Then
'       With rsRequerimiento
'            frmCargarRequerimientos.ctlTipoRequerimiento.Valor = CargarCombo(!IDTIPORECEPCION, frmCargarRequerimientos.ctlTipoRequerimiento.Valor)
'            frmCargarRequerimientos.cboTipoRequerimiento.ListIndex = CargarCombo(!IDTIPOREQUERIMIENTO, frmCargarRequerimientos.cboTipoRequerimiento)
'          Rem  frmCargarRequerimientos.cboTomo.ListIndex = CargarCombo(71, frmCargarRequerimientos.cboTomo)
'           Rem frmCargarRequerimientos.ctlClientes1.ListIndex = CargarCombo(!id_cliente, frmCargarRequerimientos.ctlClientes1)
'            frmCargarRequerimientos.TXTNumeroRequerimiento = CLng(!IDREQUERIMIENTO)
'            frmCargarRequerimientos.txtEstadoRequerimiento = CInt(!IDESTADO)
'           Rem frmCargarRequerimientos.txtClientePedido = CInt(!PEDIDOCLIENTE)
'            If Not IsNull(!Fecha) Then
'                frmCargarRequerimientos.maskHorafax = Format(CStr(!Fecha), "HH:MM")
'                frmCargarRequerimientos.maskFechafax = Format(CStr(!Fecha), "DD/MM/YYYY")
'
'            Else
'                frmCargarRequerimientos.maskHorafax = Format(CStr(!FECHARECEPCION), "HH:MM")
'                frmCargarRequerimientos.maskFechafax = Format(CStr(!FECHARECEPCION), "DD/MM/YYYY")
'            End If
'            If Not IsNull(!DESCRIPCION) Then
'                frmCargarRequerimientos.txtDescripcionCajaLibro = !DESCRIPCION
'                frmCargarRequerimientos.txtDescripcion = !DESCRIPCION
'            End If
'            If Not IsNull(!Sector) Then
'               Rem frmCargarRequerimientos.cboSector.Clear
'              Rem  frmCargarRequerimientos.cboSector.AddItem !Sector
'            End If
'
'            If Not IsNull(!SOLICITANTE) Then
'               Rem frmCargarRequerimientos.cboSolicita.Clear
'                Rem frmCargarRequerimientos.cboSolicita.AddItem !SOLICITANTE
'            End If
'            If IsNull(!FECHALIMITE) Then
'            Else
'                frmCargarRequerimientos.maskDiaLImite = Format(CStr(!FECHALIMITE), "dd/mm/yyyy")
'                frmCargarRequerimientos.maskHoraLimite = Format(CStr(!FECHALIMITE), "HH:MM")
'            End If
'            If Not IsNull(!CANTIDAD) Then
'                frmCargarRequerimientos.lblcantCajasLibros = CInt(!CANTIDAD)
'               Rem frmCargarRequerimientos.txtCantidadCajas = CInt(!CANTIDAD)
'            End If
'
'
'            frmCargarRequerimientos.lblcantCajasLibros = CStr(!CANTIDAD)
'            If Not IsNull(!Path) Then
'                        frmCargarRequerimientos.PonerImagen (!Path)
'            End If
'            Sql = " SELECT "
'            Sql = Sql & vbCrLf & " REQ.IDRequerimientos , REQ.CAJASLIBROS"
'            Sql = Sql & vbCrLf & " From"
'            Sql = Sql & vbCrLf & " REQUELIBOSCAJAS REQ"
'            Sql = Sql & vbCrLf & " Where REQ.IDRequerimientos = " & IDREQUERIMIENTO
'           frmCargarRequerimientos.fraDatosrequerimiento.Visible = True
'           frmCargarRequerimientos.fraCajas.Visible = True
'           Set rsRcajas = New ADODB.Recordset
'           rsRcajas.Open Sql, ConBasa
'           Do While Not rsRcajas.EOF
'                frmCargarRequerimientos.fraCajas.Visible = True
'                If Not IsNull(rsRcajas!CAJASLIBROS) Then
'                    CargarGrilla2 (Trim(rsRcajas!CAJASLIBROS))
'                End If
'
'                rsRcajas.MoveNext
'           Loop
'            Unload Me
'        End With
'    End If
'
'
'
'

    

End Sub

Public Function CargarCombo(dato As String, cbo As ComboBox) As Integer
Dim i As Integer
    For i = 0 To cbo.ListCount
         If CInt(dato) = Mid(cbo.List(i), 1, 3) Then
            CargarCombo = i
            Exit Function
         End If
    Next

End Function

Public Sub CargarGrilla2(Valor As String)

Dim R As Integer
Dim c As Integer

    With frmCargarRequerimientos
        For R = 1 To .grdCajasLibros.Rows - 1
            For c = 1 To .grdCajasLibros.Cols - 1
                If .grdCajasLibros.TextMatrix(R, c) = Valor Then
                    MsgBox "La Caja " & Valor & " ya esta Cargada", vbInformation
                    .txtCajaLibroDesde = ""
                    .txtCajaLibroHasta = ""
                    Exit Sub
                End If
                If .grdCajasLibros.TextMatrix(R, c) = "" Then
                    .grdCajasLibros.TextMatrix(R, c) = Valor
                    Exit Sub
                End If
            Next
        Next
        .grdCajasLibros.AddItem ""
        .grdCajasLibros.TextMatrix(.grdCajasLibros.Rows - 1, 1) = Valor
    End With
End Sub

Private Sub mskFiltro_GotFocus()
mskFiltro.SelStart = 1
End Sub
