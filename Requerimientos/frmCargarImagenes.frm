VERSION 5.00
Begin VB.Form frmCargarImagenes 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   1500
      TabIndex        =   0
      Top             =   2220
      Width           =   1935
   End
End
Attribute VB_Name = "frmCargarImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Dim CN As New ADODB.Connection
 Dim RS As New ADODB.Recordset
Dim Paso As String
Dim strCnn
     strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\REMITOS\REMITO1.mdb"

 CN.Open strCnn
  RS.Open "SELECT * FROM REMITOS", CN

Do While Not RS.EOF
    Paso = Mid(RS!Suspense_File, 20) & "F"
    FileCopy "D:\REMITOS\" & Paso, "\\Server1basa\fax\remitos\" & "REMITO " & CLng(Mid(RS!CODIGO_BARRA, 4)) & ".Tif"

    RS.MoveNext
Loop



End Sub
