VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmArchivistaPorHora 
   Caption         =   "Horas Archivista"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   6750
   Begin VB.TextBox txtRemito 
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   2820
      TabIndex        =   14
      Top             =   2580
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5460
      TabIndex        =   13
      Top             =   2580
      Width           =   1140
   End
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   375
      Left            =   1500
      TabIndex        =   12
      Top             =   480
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   661
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1500
      TabIndex        =   11
      Top             =   60
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   661
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
      Left            =   4140
      TabIndex        =   3
      Top             =   2580
      Width           =   1200
   End
   Begin MSMask.MaskEdBox mskHoraInicio 
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   1380
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      Top             =   960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskHoraFin 
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   1800
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Caption         =   "Remito:"
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
      Left            =   3000
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Diferencia:"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Hora Fin:"
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
      Left            =   180
      TabIndex        =   9
      Top             =   1860
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Hora Inicio:"
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
      Left            =   180
      TabIndex        =   8
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Dia :"
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
      Left            =   180
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Personal : "
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
      Left            =   180
      TabIndex        =   6
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label1 
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
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDiferencia 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   2475
   End
End
Attribute VB_Name = "frmArchivistaPorHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim Fecha1 As String
Dim Fecha2 As String
Dim fecha
    Fecha1 = FechaSegundoServerTipo(mskFecha.Text & " " & mskHoraInicio.Text & ":00")
    Fecha2 = FechaSegundoServerTipo(mskFecha.Text & " " & mskHoraFin.Text & ":00")
    fecha = FechaServerTipo(mskFecha.Text)
    
    

Sql = " INSERT INTO HORAS_ARCHIVISTA"
Sql = Sql & " (COD_CLIENTE, COD_PERSONAL, FECHA, HORA_INICIO,"
Sql = Sql & " HORA_FIN, DIFERENCIA, REMITO)"
Sql = Sql & " VALUES ("
Sql = Sql & ctlCliente.Valor & "," & ctlPersonal.Valor & "," & fecha & "," & Fecha1 & ","
Sql = Sql & Fecha2 & ",'" & lblDiferencia.Caption & "'," & txtRemito.Text & " )"
ExecutarSql Sql

LimpiarMask mskFecha
LimpiarMask mskHoraInicio
LimpiarMask mskHoraFin
lblDiferencia.Caption = ""
mskFecha.SetFocus

End Sub

Private Sub cmdImprimir_Click()
 Dim Sql As String
    If IsNull(ctlCliente.Valor) Then
        MsgBox "Error en cliente"
        Exit Sub
    End If
    
    If Not IsDate(mskFecha.Text) Then
        MsgBox "Error en fecha"
        Exit Sub
    End If
    
   
    Sql = " SELECT *  From V_HORASARCHIVISTA"
    Sql = Sql & " Where COD_CLIENTE =" & ctlCliente.Valor
    Sql = Sql & " AND   FECHA > " & FechaServerTipo(mskFecha.Text)
    Sql = Sql & " ORDER BY FECHA  "
    frmReportes.ImprimirReporte PasoReportes & "rptHorasArchivista.RPT", Sql, True
End Sub

Private Sub Form_Load()
ctlPersonal.TipoControl = Personal
ctlCliente.TipoControl = Cliente
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
End If

End Sub

Private Sub mskHoraFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub mskHoraFin_LostFocus()
Dim Fecha1 As String
Dim Fecha2 As String
    Fecha1 = mskFecha.Text & " " & mskHoraInicio.Text
    Fecha2 = mskFecha.Text & " " & mskHoraFin.Text
    lblDiferencia = Format(DateDiff("n", Fecha1, Fecha2) / 60, "##.00")
End Sub

Private Sub mskHoraInicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub
