VERSION 5.00
Begin VB.Form frmImagenesWeb 
   Caption         =   "Imgenes web"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form4"
   ScaleHeight     =   5250
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEtiqueta 
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Text            =   "120006230107"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdActualizarDigital 
      Caption         =   "Actualizar digital"
      Height          =   615
      Left            =   1260
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "frmImagenesWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strConAsp As String



Private Sub Command1_Click()

End Sub

Private Sub cmdActualizarDigital_Click()
    Dim Paso As String
    Paso = "C:\Archivos_Digitales\PDF\ConsultasDigitales\"
    Dim sql As String
    
    Dim rs As New ADODB.Recordset
    Dim ConAsp As New ADODB.Connection
    

sql = " SELECT     id, codigo"
sql = sql & " From basa.dbo.elementos "
sql = sql & " WHERE codigo = '" & txtEtiqueta.Text & "'"

rs.Open sql, strConAsp
ConAsp.Open strConAsp
    If Not rs.EOF Then
        
        sql = "Update referencia"
        sql = sql & " SET pathLegajo ='" & Paso & txtEtiqueta.Text & ".PDF'"
        sql = sql & "  Where referencia.elemento_id =" & rs!ID
        ConAsp.Execute sql
         MsgBox "teRMINADO"
    End If
    



End Sub

Private Sub Form_Load()

 
    strConAsp = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
    

End Sub
