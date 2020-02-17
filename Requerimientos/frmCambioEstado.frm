VERSION 5.00
Begin VB.Form frmCambioEstado 
   Caption         =   "CAJAS / LIBROS ENCONTRADOS"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   960
   ScaleWidth      =   4920
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      Picture         =   "frmCambioEstado.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   0
      Width           =   915
   End
   Begin VB.TextBox txtLecturarequerimiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
End
Attribute VB_Name = "frmCambioEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_Change()



End Sub

Private Sub txtLecturarequerimiento_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
            If UCase(Mid(txtLecturarequerimiento.Text, 1, 3)) = "D10" Then
                Dim RS As ADODB.Recordset
                Dim rs2 As ADODB.Recordset
                Set RS = New ADODB.Recordset
                RS.Open "Select * from REQUERIMIENTO WHERE IDESTADO = 3 AND IDREQUERIMIENTO = " & CLng(Mid(txtLecturarequerimiento.Text, 4)), conbasa
                If Not RS.EOF Then
                   CRequerimientos.Clear
                   CRequerimientos.Add Format(CStr(RS!FECHARECEPCION), "dd/mm/yyyy"), 3, CLng(Mid(txtLecturarequerimiento.Text, 4)), CInt(RS!IDTipoRequerimiento)
                   Set rs2 = New ADODB.Recordset
                   rs2.Open "Select * from personal where jefeplanta = 1", conbasa
                   EstadoFinal = 4
                   If Not rs2.EOF Then
                        CRequerimientos.CambioEstado CInt(rs2!idPersonal), True, 3, 4
                   Else
                        CRequerimientos.CambioEstado 99, True
                   End If
                
                Else
                    MsgBox "YA ESTA CARGADO"
                End If
            Else
                MsgBox "EL FORMULARIO NOES UN REQUERIMIENTO VALIDO"
            End If
            txtLecturarequerimiento.Text = ""
End If


End Sub
