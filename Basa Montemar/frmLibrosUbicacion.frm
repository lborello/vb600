VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmLibrosUbicacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros Ubicacion"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9435
   Begin VB.Frame Frame1 
      Caption         =   "Generar Rotulos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   3900
      Width           =   7755
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
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
         Left            =   6300
         TabIndex        =   12
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblFinal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4560
         TabIndex        =   16
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblInicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "Fin:"
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
         Left            =   4140
         TabIndex        =   14
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Inicio:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2100
         TabIndex        =   13
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdRotuloChico 
      Caption         =   "Rotulo"
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
      Left            =   8040
      TabIndex        =   7
      Top             =   4260
      Width           =   1200
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
      Height          =   330
      Left            =   8040
      TabIndex        =   0
      Top             =   3840
      Width           =   1200
   End
   Begin MSMask.MaskEdBox mskDesde 
      Height          =   315
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid grdUbicacion 
      Height          =   3135
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskHasta 
      Height          =   315
      Left            =   8340
      TabIndex        =   3
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####"
      PromptChar      =   "_"
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   1020
      TabIndex        =   8
      Top             =   120
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
   End
   Begin VB.Label Label5 
      Caption         =   "Libro Desde :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5340
      TabIndex        =   6
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   180
      Width           =   675
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmLibrosUbicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdImprimir_Click()

End Sub

Private Sub cmdGenerar_Click()
    Dim rsLibrosNuevos As New ADODB.Recordset
    Dim sSQL As String
    Dim MAX_NRO_LIBRO As Long
    Dim MAX_NRO_LIBRO_INTERNO As Long
    Dim Cliente As Integer
    Cliente = ctlCliente.Valor
    Dim i As Integer
    
    sSQL = " SELECT MAX(NRO_LIBRO) AS MAX_NRO_LIBRO"
    sSQL = sSQL & vbCrLf & " FROM LIBROS "
    rsLibrosNuevos.Open sSQL, ConActiva, 0, 1
    MAX_NRO_LIBRO = rsLibrosNuevos!MAX_NRO_LIBRO + 1
        
    Set rsLibrosNuevos = New ADODB.Recordset
    sSQL = "  SELECT MAX(NRO_LIBRO_INTERNO) AS MAX_NRO_LIBRO_INTERNO"
    sSQL = sSQL & vbCrLf & " From libros Where COD_CLIENTE =" & Cliente
    rsLibrosNuevos.Open sSQL, ConActiva, 0, 1
    If Not rsLibrosNuevos.EOF Then
       If IsNull(rsLibrosNuevos!MAX_NRO_LIBRO_INTERNO) Then
            MAX_NRO_LIBRO_INTERNO = 1
       Else
            MAX_NRO_LIBRO_INTERNO = rsLibrosNuevos!MAX_NRO_LIBRO_INTERNO + 1
       End If
    Else
       MAX_NRO_LIBRO_INTERNO = 1
    End If
    lblInicio.Caption = MAX_NRO_LIBRO_INTERNO
    lblInicio.Refresh
    For i = 1 To txtCantidad.Text
        sSQL = " INSERT INTO LIBROS (NRO_LIBRO, NRO_LIBRO_INTERNO, COD_CLIENTE, ESTADO) "
        sSQL = sSQL & vbCrLf & " VALUES (" & MAX_NRO_LIBRO & "," & MAX_NRO_LIBRO_INTERNO & "," & Cliente & ",5)"
         ExecutarSql (sSQL)
        MAX_NRO_LIBRO = MAX_NRO_LIBRO + 1
        MAX_NRO_LIBRO_INTERNO = MAX_NRO_LIBRO_INTERNO + 1
    Next
    lblFinal.Caption = MAX_NRO_LIBRO_INTERNO
    lblFinal.Refresh
    MsgBox "Terminado"
 End Sub

Private Sub cmdLimpiar_Click()
    grdUbicacion.Clear
    grdUbicacion.Rows = 2
    TitulosGrillasUbicacion
End Sub

Private Sub cmdRotuloChico_Click()
    Dim sSQL As String
    Dim i As Integer
     sSQL = " SELECT * "
        sSQL = sSQL & " From libros"
        sSQL = sSQL & " Where  COD_CLIENTE = " & ctlCliente.Valor
        sSQL = sSQL & " AND NRO_LIBRO_INTERNO in (0 "
    With grdUbicacion
        For i = 1 To .Rows - 2
          sSQL = sSQL & "," & .TextMatrix(i, 2)
        Next
    End With
     sSQL = sSQL & " ) Order by NRO_LIBRO_INTERNO "
     frmReportes.ImprimirReporte PasoReportes & "rptRotulosLibros.rpt", sSQL, True
End Sub


Private Sub Command1_Click()
   
End Sub

Private Sub Form_Load()
     TitulosGrillasUbicacion
     ctlCliente.TipoControl = Cliente
End Sub

Public Sub TitulosGrillasUbicacion()
     
     With grdUbicacion
        .Cols = 5
        .ColWidth(0) = 100
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColAlignment(0) = flexAlignCenterCenter ' flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "Razon Social"
        .TextMatrix(0, 2) = "Libro_interno"
        .TextMatrix(0, 3) = "Estado"
        .TextMatrix(0, 4) = "Ubicación"
        
    End With
End Sub

Private Sub mskDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       InsertarGrilla mskDesde.Text, ctlCliente.Valor
       mskDesde = ""
    End If
End Sub

Private Sub mskHasta_KeyPress(KeyAscii As Integer)
    Dim i As Long
        If KeyAscii = 13 Then
           MousePointer = 13
               For i = mskDesde To mskHasta
                   InsertarGrilla i, ctlCliente.Valor
               Next
           mskDesde = ""
           mskHasta = ""
           MousePointer = 0
        End If
End Sub


Public Sub InsertarGrilla(NRO_LIBRO_INTERNO As Long, COD_CLIENTE As Integer)
    Dim rsLibros As New ADODB.Recordset
    Dim R As Integer
    Dim sSQL As String
    
    sSQL = "  SELECT NRO_LIBRO, NRO_LIBRO_INTERNO,COD_CLIENTE ,UBICACION, Razon_Social ,estado"
    sSQL = sSQL & vbCrLf & "   From libros, CLIENTES"
    sSQL = sSQL & vbCrLf & "  WHERE LIBROS.COD_CLIENTE = CLIENTES.ID_CLIENTE AND"
    sSQL = sSQL & vbCrLf & " (COD_CLIENTE =" & ctlCliente.Valor & " ) AND (NRO_LIBRO_INTERNO = " & NRO_LIBRO_INTERNO & ")"
 
    rsLibros.Open sSQL, ConActiva, 0, 1
    Do While Not rsLibros.EOF
         With grdUbicacion
             R = .Rows - 1
             .TextMatrix(R, 0) = rsLibros.Fields("Cod_Cliente").value
             .TextMatrix(R, 1) = Trim(UCase(rsLibros.Fields("RAZON_SOCIAL").value))
             .TextMatrix(R, 2) = rsLibros.Fields("NRO_LIBRO_INTERNO").value
             Select Case rsLibros.Fields("Estado").value
             Case "2"
                 .TextMatrix(R, 3) = "Planta"
             Case "3"
                 .TextMatrix(R, 3) = "Consulta"
             Case "4"
                 .TextMatrix(R, 3) = "Reserva"
             Case "5"
                 .TextMatrix(R, 3) = "Cliente"
             End Select
             If IsNull(rsLibros.Fields("UBICACION").value) Then
                 .TextMatrix(R, 4) = ""
             Else
                 .TextMatrix(R, 4) = rsLibros.Fields("UBICACION").value
             End If
        End With
        grdUbicacion.AddItem ""
        rsLibros.MoveNext
    Loop
End Sub

