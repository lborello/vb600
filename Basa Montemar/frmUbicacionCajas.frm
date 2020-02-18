VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmCajasUbicacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UBICACIÓN"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   10875
   Begin Controles.cltGenerico cltGenerico1 
      Height          =   375
      Left            =   420
      TabIndex        =   13
      Top             =   4560
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   661
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdLimpiarOrden 
      Caption         =   "Limpiar Orden"
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
      Left            =   7860
      TabIndex        =   11
      Top             =   840
      Width           =   1515
   End
   Begin VB.TextBox txtOrden 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   840
      Width           =   4695
   End
   Begin VB.ComboBox cboOrden 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmUbicacionCajas.frx":0000
      Left            =   1320
      List            =   "frmUbicacionCajas.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   1515
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   4620
      Width           =   1515
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   8460
      TabIndex        =   4
      Top             =   4620
      Width           =   1515
   End
   Begin MSMask.MaskEdBox mskDesde 
      Height          =   435
      Left            =   6300
      TabIndex        =   1
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#######"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid grdUbicacion 
      Height          =   2955
      Left            =   180
      TabIndex        =   0
      Top             =   1440
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   9
   End
   Begin MSMask.MaskEdBox mskHasta 
      Height          =   435
      Left            =   9180
      TabIndex        =   2
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#######"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "Orden:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8220
      TabIndex        =   6
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Cajas Desde :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
      End
   End
End
Attribute VB_Name = "frmCajasUbicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOrden_Click()
    TxtOrden.Text = TxtOrden.Text & cboOrden.Text & " ASC ,"
End Sub

Private Sub cmdImprimir_Click()
    Dim i As Integer
    Dim Filtro As String
    Dim Sql As String
       If IsNull(ctlCliente.Valor) Then
            MsgBox "Ingrese el Numero de cliente", vbInformation
            Exit Sub
       End If
        For i = 1 To grdUbicacion.Rows - 1
            grdUbicacion.Row = i
            Filtro = Filtro & grdUbicacion.TextMatrix(i, 2) & ","
        Next
        Sql = " SELECT *  FROM  CONTENEDOR"
        Sql = Sql & " where cod_cliente = " & ctlCliente.Valor
        Sql = Sql & "  and nro_caja in( " & Mid(Filtro, 1, Len(Filtro) - 2) & " )"
        If TxtOrden.Text = "" Then
            Sql = Sql & "  ORDER BY ESTANTERIA, HORIZONTAL,VERTICAL "
        Else
            Sql = Sql & Mid(TxtOrden.Text, 1, Len(TxtOrden.Text) - 1)
        End If
        frmReportes.ImprimirReporte PasoReportes & "rptBuscarCajas.rpt", Sql, True
       
End Sub
Private Sub cmdLimpiar_Click()
    grdUbicacion.Clear
    grdUbicacion.Rows = 2
    TitulosGrillasUbicacion
    TxtOrden.Text = ""
End Sub

Private Sub cmdLimpiarOrden_Click()
TxtOrden.Text = ""
End Sub

Private Sub Form_Load()
     TitulosGrillasUbicacion
     ctlCliente.TipoControl = Cliente
End Sub

Public Sub TitulosGrillasUbicacion()
     With grdUbicacion
        .ColWidth(0) = 100
        .ColWidth(1) = 2000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 2600
        
        .ColAlignment(0) = flexAlignCenterCenter ' flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        
        .TextMatrix(0, 1) = "Razon Social"
        .TextMatrix(0, 2) = "Nº Caja"
        .TextMatrix(0, 3) = "Estanteria"
        .TextMatrix(0, 4) = "Horizontal"
        .TextMatrix(0, 5) = "Vertical"
        .TextMatrix(0, 6) = "Adl/Atras"
        .TextMatrix(0, 7) = "Estado"
        .TextMatrix(0, 8) = "Ub/Prov."
    End With
End Sub

Private Sub grdUbicacion_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnu
End If
End Sub

Private Sub mnuCopiar_Click()
     Dim i As Integer
     Dim R As Integer
Dim DATO As String
     
For i = 0 To grdUbicacion.Cols - 1
    DATO = DATO & vbTab & grdUbicacion.TextMatrix(0, i)
Next
    DATO = DATO & vbCrLf
For R = 1 To grdUbicacion.Rows - 1
    For i = 0 To grdUbicacion.Cols - 1
        DATO = DATO & vbTab & grdUbicacion.TextMatrix(R, i)
    Next
    DATO = DATO & vbCrLf
Next

Clipboard.Clear
Clipboard.SetText DATO
MsgBox "Copia Terminada"
End Sub

Private Sub mskDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       InsertarGrilla mskDesde.Text, ctlCliente.Valor, True
       mskDesde = ""
    End If
End Sub

Private Sub mskHasta_KeyPress(KeyAscii As Integer)
    Dim i As Long
        If KeyAscii = 13 Then
           MousePointer = 13
               For i = mskDesde To mskHasta
                   InsertarGrilla i, ctlCliente.Valor, True
               Next
           mskDesde = ""
           mskHasta = ""
           MousePointer = 0
        End If
End Sub


Public Sub InsertarGrilla(NRO_CAJA As Long, COD_CLIENTE As Integer, Mensaje As Boolean)
    Dim rsContenedor As New ADODB.Recordset
    Dim R As Integer
    Dim sSQL As String
    
'    If CajaRepetida(NRO_CAJA) Then
'        If Mensaje Then
'            MsgBox "Caja Repetida", vbInformation
'            Exit Sub
'        Else
'            Exit Sub
'        End If
'    End If
    
    
    
    
    
    sSQL = "   SELECT CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL,"
    sSQL = sSQL & vbCrLf & " CONTENEDOR.VERTICAL, CONTENEDOR.ADELANTE_ATRAS,"
    sSQL = sSQL & vbCrLf & " CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE,"
    sSQL = sSQL & vbCrLf & " CONTENEDOR.NRO_CAJA, CLIENTES.RAZON_SOCIAL,CONTENEDOR.UB_PROVISORIA"
    sSQL = sSQL & vbCrLf & " From CONTENEDOR, CLIENTES"
    sSQL = sSQL & vbCrLf & "  WHERE CONTENEDOR.COD_CLIENTE = CLIENTES.ID_CLIENTE AND"
    sSQL = sSQL & vbCrLf & " (CONTENEDOR.COD_CLIENTE =" & COD_CLIENTE & " ) AND (CONTENEDOR.NRO_CAJA = " & NRO_CAJA & ")"
    rsContenedor.Open sSQL, ConActiva, 0, 1
    Do While Not rsContenedor.EOF
         With grdUbicacion
             R = .Rows - 1
             .TextMatrix(R, 0) = rsContenedor.Fields("Cod_Cliente").value
             .TextMatrix(R, 1) = Trim(UCase(rsContenedor.Fields("RAZON_SOCIAL").value))
             .TextMatrix(R, 2) = rsContenedor.Fields("NRO_CAJA").value
             .TextMatrix(R, 3) = rsContenedor.Fields("Estanteria").value
             .TextMatrix(R, 4) = rsContenedor.Fields("Horizontal").value
             .TextMatrix(R, 5) = rsContenedor.Fields("Vertical").value
             Select Case rsContenedor.Fields("Adelante_Atras").value
             Case "1"
                 .TextMatrix(R, 6) = "Atras"
             Case "2"
                 .TextMatrix(R, 6) = "Frente"
             End Select
             Select Case rsContenedor.Fields("Estado").value
             Case "2"
                 .TextMatrix(R, 7) = "Planta"
             Case "3"
                 .TextMatrix(R, 7) = "Consulta"
             Case "4"
                 .TextMatrix(R, 7) = "Reserva"
             Case "5"
                 .TextMatrix(R, 7) = "Cliente"
             End Select
             If IsNull(rsContenedor.Fields("UB_PROVISORIA").value) Then
                 .TextMatrix(R, 8) = ""
             Else
                 .TextMatrix(R, 8) = rsContenedor.Fields("UB_PROVISORIA").value
             End If
        End With
        grdUbicacion.AddItem ""
        rsContenedor.MoveNext
    Loop
End Sub

Public Function CajaRepetida(Caja As Long) As Boolean
    Dim R  As Integer
    CajaRepetida = False
    For R = 1 To grdUbicacion.Rows - 1
       If grdUbicacion.TextMatrix(R, 2) <> "" Then
           If grdUbicacion.TextMatrix(R, 2) = Caja Then
                CajaRepetida = True
                Exit Function
           End If
        End If
        
        
    Next

End Function
