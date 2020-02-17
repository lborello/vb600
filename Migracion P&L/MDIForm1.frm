VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8955
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11685
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MNUCAJA 
      Caption         =   "CAJA"
   End
   Begin VB.Menu mnuLectura 
      Caption         =   "Lectura"
   End
   Begin VB.Menu Elelmento 
      Caption         =   "Elemento"
   End
   Begin VB.Menu MnuImagenes 
      Caption         =   "Imagenes"
   End
   Begin VB.Menu frmControlReferencias 
      Caption         =   "control de referencias"
   End
   Begin VB.Menu mnuMigracion 
      Caption         =   "migracion"
   End
   Begin VB.Menu mnuCambioLectura 
      Caption         =   "Cambio de Lectura"
   End
   Begin VB.Menu controlreferenciasasp 
      Caption         =   "controlreferenciasasp"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub controlreferenciasasp_Click()
Form4.Show
Form4.SetFocus

End Sub

Private Sub Elelmento_Click()
Form3.Show
Form3.SetFocus
End Sub

Private Sub frmControlReferencias_Click()
frmContronReferencias.Show
frmContronReferencias.SetFocus
End Sub

Private Sub MNUCAJA_Click()
frmGrilla.Show
frmGrilla.SetFocus
End Sub

Private Sub mnuCambioLectura_Click()
    frmCambiosporLectura.Show
    frmCambiosporLectura.SetFocus
End Sub

Private Sub MnuImagenes_Click()
    frmImagenesWeb.Show
    frmImagenesWeb.SetFocus
End Sub

Private Sub mnuLectura_Click()
    frmLecturas.Show
    frmLecturas.SetFocus
End Sub

Private Sub mnuMigracion_Click()
    frmMigracion.Show
    frmMigracion.SetFocus

End Sub
