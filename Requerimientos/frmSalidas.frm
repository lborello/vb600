VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{40CE97D1-1C1F-47E7-B2C4-A9B643CAAFFD}#5.0#0"; "Controles.ocx"
Begin VB.Form frmSalidas 
   Caption         =   "Salidas"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   6195
   Begin Controles.ViewImg ViewImg1 
      Height          =   3555
      Left            =   60
      TabIndex        =   4
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6271
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2220
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox txtRequerimientos 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   60
      Width           =   1875
   End
   Begin MSFlexGridLib.MSFlexGrid grdRequerimientos 
      Height          =   2715
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4789
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
 Dim Sql As String
Dim i As Integer

For i = 1 To grdRequerimientos.Rows - 1
   If grdRequerimientos.TextMatrix(i, 1) <> "" Then
        Sql = " Update Requerimiento Set IDESTADO = 6 Where (IDESTADO =6) And (IDREQUERIMIENTO =" & grdRequerimientos.TextMatrix(i, 1) & ")"
        Debug.Print Sql
        ' conbasa.Execute
    End If
 Next
    grdRequerimientos.Clear
    grdRequerimientos.Rows = 1
    grdRequerimientos.TextMatrix(0, 1) = "Requerimientos"
    grdRequerimientos.ColWidth(0) = 100
    grdRequerimientos.ColWidth(1) = 3000
    grdRequerimientos.Redraw = True
    grdRequerimientos.Refresh
End Sub

Private Sub Form_Load()
    grdRequerimientos.Clear
    grdRequerimientos.Rows = 1
    grdRequerimientos.TextMatrix(0, 1) = "Requerimientos"
    grdRequerimientos.ColWidth(0) = 100
    grdRequerimientos.ColWidth(1) = 3000
    grdRequerimientos.Redraw = True
    grdRequerimientos.Refresh
End Sub

Private Sub txtRequerimientos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdRequerimientos.AddItem vbTab & txtRequerimientos
        txtRequerimientos.Text = ""
    End If
End Sub
