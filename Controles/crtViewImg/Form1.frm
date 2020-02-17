VERSION 5.00
Begin VB.Form frm1 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ViewImg ViewImg1 
      Height          =   4335
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7646
   End
   Begin VB.PictureBox handle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   0
      Left            =   0
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim viewScalaX As Integer
Dim viewScalaY As Integer


Private Sub Form_Load()
    viewScalaX = Me.Width
    viewScalaY = Me.Height
    
    ViewImg1.MostrarImagen ("c:\Burgos Zapata Juan  Horacio.tif")
End Sub


Private Sub Form_Resize()

If ViewImg1.sizeC(CDbl(viewScalaX), CDbl(viewScalaY), CDbl(Me.Width), CDbl(Me.Height)) Then
    viewScalaX = Me.Width
    viewScalaY = Me.Height

End If


End Sub

