VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmCargarFechaHora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha Hora"
   ClientHeight    =   9540
   ClientLeft      =   1305
   ClientTop       =   1290
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   10830
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtBorrar 
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grdRemitos 
      Height          =   7455
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13150
      _Version        =   393216
   End
   Begin VB.TextBox txtLote 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdRequerimiento 
      Caption         =   "Aceptar"
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
      Left            =   9480
      TabIndex        =   8
      Top             =   9000
      Width           =   1035
   End
   Begin VB.TextBox txtRequerimientoE 
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
      IMEMode         =   3  'DISABLE
      Left            =   7920
      TabIndex        =   6
      Top             =   840
      Width           =   2595
   End
   Begin VB.TextBox txtRequerimiento 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   3075
   End
   Begin Controles.cltGenerico cltGenerico1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   8640
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   8640
      Width           =   1035
   End
   Begin VB.TextBox txtRemito 
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2355
   End
   Begin VB.Label Label4 
      Caption         =   "Borrar "
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Lote"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Requerimiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Remito:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmCargarFechaHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FECHARE
Dim IDREQUER As Long
Dim FECHARECEPCION As Date
Dim IDREMITO As Long
Dim OrdenRemito As Integer
Private Sub cmdAceptar_Click()
    Dim sql As String
    Dim conRemito As New ADODB.Connection
    Dim i As Integer
    
    
    On Error GoTo salir

If Trim(txtLote.Text) <> "" Then

        For i = 1 To grdRemitos.Rows - 1
            If IsNumeric(grdRemitos.TextMatrix(i, 1)) Then
        
                sql = " Update REMITOS_CUERPO SET "
                sql = sql & " FECHA_LECTURA_REMITO = '" & txtLote.Text & "  Orden:" & grdRemitos.TextMatrix(i, 0) & "'"
                sql = sql & " Where NRO_REMITO = " & grdRemitos.TextMatrix(i, 1)
                If ExecutarSql(sql) <> 1 Then
                    MsgBox "No se registro el remito " & grdRemitos.TextMatrix(i, 1)
                End If
                sql = " Update Requerimiento "
                sql = sql & " SET  IDESTADO = 7  "
                sql = sql & " , FECHA_LECTURA= '" & txtLote.Text & "  Orden:" & grdRemitos.TextMatrix(i, 0) & "'"
                sql = sql & " Where IDREMITO = " & grdRemitos.TextMatrix(i, 1)
                sql = sql & " AND IDESTADO > 5 "
                ExecutarSql sql
                
            Else
        
                sql = " Update REMITOS_CUERPO SET "
                sql = sql & " FECHA_LECTURA_REMITO = '" & txtLote.Text & "  Orden:" & grdRemitos.TextMatrix(i, 0) & "'"
                sql = sql & " Where  NRO_REMITO > 178851 AND   NRO_REM_PROV = '" & Trim(grdRemitos.TextMatrix(i, 1)) & "'"
                If ExecutarSql(sql) < 1 Then
                    MsgBox "No se registro el remito " & grdRemitos.TextMatrix(i, 1)
                End If
        
        
            End If
        Next
Else
    MsgBox "Ingrese el lote"
End If




Unload Me
salir:
MsgBox Err.Description

End Sub

Private Sub cmdBorrar_Click()
 Dim i As Integer
 
 For i = 1 To grdRemitos.Rows - 1
    If Trim(txtBorrar.Text) = grdRemitos.TextMatrix(i, 0) Then
        grdRemitos.RemoveItem i
        Exit For
    End If
 
 
 Next
 
 
 For i = 1 To grdRemitos.Rows - 1
  grdRemitos.TextMatrix(i, 0) = i
 
 
 Next
 
OrdenRemito = i - 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub maskFecha_GotFocus()

End Sub

Private Sub cmdRequerimiento_Click()
   
    Dim sql As String
On Error GoTo salir

    sql = " Update Requerimiento "
    sql = sql & " SET  IDESTADO =7 "
    sql = sql & " , FECHA_LECTURA= '" & SysDate_DD_MM_YYYY & "'"
    sql = sql & " Where IDREQUERIMIENTO in (  " & Mid(txtRequerimiento.Text, 2) & ")"
    sql = sql & " AND IDESTADO > 5 "
    ExecutarSql sql
    Rem  txtRemitos.Text = ""
    Unload Me
salir:
    MsgBox Err.Description
    
End Sub


Private Sub Form_Load()
OrdenRemito = 0
grdRemitos.TextMatrix(0, 0) = "Orden"
grdRemitos.TextMatrix(0, 1) = "Remito"
grdRemitos.ColWidth(0) = 500
grdRemitos.ColWidth(1) = 2000

End Sub

Private Sub txtRemito_KeyPress(KeyAscii As Integer)
     Dim Remito As String
     If KeyAscii = 13 Then
     
     If Trim(txtRemito.Text) = "" Then
      Exit Sub
     End If
     
     Rem luisok
     
          If IsNumeric(txtRemito.Text) Then
            If Len(txtRemito.Text) = 5 Then
                Remito = "0001-000" & txtRemito.Text
            End If
            
            If Len(txtRemito.Text) > 5 Then
                 Remito = txtRemito.Text
            End If
            If Len(Trim(txtRemito.Text)) = 12 Then
              Remito = "0001-" & Mid(Trim(txtRemito.Text), 5)
            End If
            
          Else
            If Len(Trim(txtRemito.Text)) = 13 Then
              Remito = Trim(txtRemito.Text)
            End If
            
            If UCase(Mid(txtRemito.Text, 1, 2)) = "RM" Then
              Remito = "0001-000" & Mid(txtRemito.Text, 3)
            End If
          End If
         If Remito = "" Then
            MsgBox "Error Remito"
            Exit Sub
         End If
         
          
        OrdenRemito = OrdenRemito + 1
        If OrdenRemito = 1 Then
          grdRemitos.TextMatrix(1, 0) = 1
          grdRemitos.TextMatrix(1, 1) = Remito
        Else
          grdRemitos.AddItem Trim(OrdenRemito & vbTab & Remito)
        End If
        txtRemito.Text = ""
        grdRemitos.ScrollTrack = True
        If OrdenRemito > 30 Then
            grdRemitos.TopRow = OrdenRemito - 10
        End If
    End If
 End Sub

Private Sub txtRequerimientoE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
   
       
       txtRequerimiento.Text = "," & CLng(txtRequerimientoE.Text) & vbCrLf & txtRequerimiento.Text
       txtRequerimientoE.Text = ""
     
   End If
   
End Sub
