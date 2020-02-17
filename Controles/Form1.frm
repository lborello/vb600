VERSION 5.00
Object = "*\AControles.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin Controles.cltGenerico cltGenerico1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      _extentx        =   5953
      _extenty        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cltGenerico1.TipoControl = Cliente
End Sub
