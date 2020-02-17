VERSION 5.00
Begin VB.Form frmEstadoRequerimientoTerminado 
   Caption         =   "Estado Requerimiento"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   3465
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
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
      Left            =   600
      TabIndex        =   5
      Top             =   5220
      Width           =   2235
   End
   Begin VB.TextBox txtRequerimientos 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1020
      Width           =   3135
   End
   Begin VB.TextBox txtLectura 
      BackColor       =   &H00C0FFC0&
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
      Left            =   1620
      TabIndex        =   2
      Top             =   540
      Width           =   1575
   End
   Begin VB.TextBox txtHojadeRuta 
      BackColor       =   &H00C0E0FF&
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
      Left            =   1620
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Requerimiento"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Hoja de Ruta"
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
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmEstadoRequerimientoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
    Dim sql As String
    
    
    sql = " UPDATE    REQUERIMIENTO"
    sql = sql & " SET    COD_HOJA_RUTA_TERMINADO =0"
    sql = sql & " WHERE  COD_HOJA_RUTA_TERMINADO =" & txtHojadeRuta.Text
    ExecutarSql sql
    
     
    sql = " UPDATE    REQUERIMIENTO"
    sql = sql & " SET    COD_HOJA_RUTA_TERMINADO =" & txtHojadeRuta.Text
    sql = sql & " WHERE  (IDREQUERIMIENTO IN (" & Mid(Replace(Trim(txtRequerimientos.Text), vbCrLf, ""), 2) & "))"
    ExecutarSql sql
    
    
    sql = " UPDATE    REQUERIMIENTO"
    sql = sql & " SET    IDESTADO =6"
    sql = sql & " WHERE  (IDESTADO < 6) AND (IDREQUERIMIENTO IN (" & Mid(Replace(Trim(txtRequerimientos.Text), vbCrLf, ""), 2) & "))"
    ExecutarSql sql
    
    
    sql = " Update HOJA_RUTA_CUERPO Set ESTADO = 100 "
    sql = sql & "  Where ID_HOJA_RUTA = " & txtHojadeRuta.Text
    ExecutarSql sql

    Unload Me
End Sub

Private Sub txtLectura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
             txtRequerimientos.Text = txtRequerimientos.Text & " ," & CLng(txtLectura) & vbCrLf
            txtLectura.Text = ""
    End If

End Sub
