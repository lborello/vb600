VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBloque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Bloque"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
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
      Left            =   5820
      TabIndex        =   11
      Top             =   4860
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
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
      Left            =   180
      TabIndex        =   8
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
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
      Left            =   4380
      TabIndex        =   5
      Top             =   4860
      Width           =   1320
   End
   Begin VB.TextBox txtUbicacion 
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
      Left            =   1560
      TabIndex        =   4
      Top             =   420
      Width           =   5595
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "Asignar"
      Enabled         =   0   'False
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
      Left            =   180
      TabIndex        =   3
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton cmdLectura 
      Caption         =   "Lectura"
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
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdBloque 
      Height          =   3615
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6376
      _Version        =   393216
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
   Begin VB.Label lblCantidad 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2340
      TabIndex        =   10
      Top             =   780
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblFechaLectura 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5700
      TabIndex        =   7
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label lblLectura 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1560
      TabIndex        =   6
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblLecturaDescripcion 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   2
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "frmBloque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsignar_Click()
    Dim i As Integer
    For i = 1 To grdBloque.Rows - 1
        grdBloque.TextMatrix(i, 3) = txtUbicacion.Text & " - " & grdBloque.TextMatrix(i, 0)
    Next
    cmdGuardar.Enabled = True
End Sub

Private Sub cmdGuardar_Click()
    Dim i As Integer
    Dim sSQL As String
        MousePointer = 11
        With grdBloque
            For i = 1 To grdBloque.Rows - 1
                sSQL = "UPDATE CONTENEDOR SET UB_PROVISORIA ='" & Trim(.TextMatrix(i, 3)) & "'"
                sSQL = sSQL & vbCrLf & " Where COD_CLIENTE = " & .TextMatrix(i, 1) & " And NRO_CAJA = " & .TextMatrix(i, 2)
               ExecutarSql sSQL
            Next
        End With
        TituloGrilla
        MousePointer = 0
End Sub

Private Sub cmdLectura_Click()
    Dim rsLectura As ADODB.Recordset
    Dim Lectura As String
    Dim sSQL As String
    
    Lectura = InputBox("Ingrese el numero de lectura", "Lectura")
    Set rsLectura = New ADODB.Recordset
    sSQL = " SELECT NUMERO_LECTURA, DESCRIPCION,FECHA_CREACION"
    sSQL = sSQL & " From LECTURA_COLECTOR_CUERPO"
    sSQL = sSQL & " Where NUMERO_LECTURA IN( " & Lectura & ")"
    rsLectura.Open sSQL, ConActiva, 0, 1
    If Not rsLectura.EOF Then
        lblFechaLectura = rsLectura!FECHA_CREACION
        lblLecturaDescripcion = rsLectura!Descripcion
        lblLectura = rsLectura!NUMERO_LECTURA
    End If
    Set rsLectura = New ADODB.Recordset
    sSQL = " SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN"
    sSQL = sSQL & " From LECTURACOLECTOR"
    sSQL = sSQL & " Where NUMERO_LECTURA IN(" & Lectura & ")"
    sSQL = sSQL & " ORDER BY ORDEN"
    rsLectura.Open sSQL, ConActiva, 0, 1
    With rsLectura
        Do While Not .EOF
            grdBloque.AddItem !Orden & vbTab & !Cliente & vbTab & !Caja
            rsLectura.MoveNext
        Loop
    End With
    lblCantidad.Caption = grdBloque.Rows - 1
    cmdAsignar.Enabled = True
End Sub

Private Sub cmdLimpiar_Click()
    TituloGrilla
End Sub

Private Sub Form_Load()
    TituloGrilla
End Sub

Public Sub TituloGrilla()
   With grdBloque
        .Clear
        .Rows = 1
        .Cols = 4
        .ColWidth(0) = 600
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 2800
        .TextArray(0) = "Orden"
        .TextArray(1) = "Cliente"
        .TextArray(2) = "Caja"
        .TextArray(3) = "Ubicación"
        lblFechaLectura.Caption = ""
        lblLectura.Caption = ""
        lblLecturaDescripcion.Caption = ""
        lblCantidad.Caption = ""
        txtUbicacion.Text = ""
        cmdGuardar.Enabled = False
        cmdAsignar.Enabled = False
    End With
End Sub

Private Sub Label2_Click()

End Sub
