VERSION 5.00
Object = "{A5EDEDF4-2BBC-45F3-822B-E60C278A1A79}#11.0#0"; "MDIVWCTL.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ViewImg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   4830
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3930
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":1D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":215C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":965E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":9F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":A812
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":AC6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":10F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":11358
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":118F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":11D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":11EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":121C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barImg 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ORC"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desplazar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Seleccionar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Texto"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Imagen"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Rotar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1560
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":12BD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":13024
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":18846
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":18C98
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":1EF32
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":26434
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":26D0E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MODICtl.MiDocView MiDocView1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      _cx             =   8493
      _cy             =   5741
      ActionState     =   0
      DocViewMode     =   0
      FitMode         =   0
      FileName        =   ""
   End
End
Attribute VB_Name = "ViewImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim doc As New MODI.Document
Dim img As New MODI.Image
Private mblnCancel As Boolean
Dim Xx As Integer
Dim yy As Integer

Public Sub MostrarImagen(sImagen As String)
   On Error GoTo mensjError
    doc.Create (sImagen)
    
    MiDocView1.Document = doc
    StatusBar1.Panels(1) = sImagen
    StatusBar1.Panels(2) = "Paginas" & MiDocView1.NumPages
    Exit Sub
mensjError:
  
End Sub

Private Sub barImg_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
     Screen.MousePointer = vbHourglass
        doc.OCR (miLANG_SPANISH)
        Screen.MousePointer = vbDefault
    Case 3
        MiDocView1.ActionState = miASTATE_PAN
        
        barImg.Buttons.Item(5).Enabled = True
        
    Case 4
        MiDocView1.ActionState = miASTATE_ZOOM
        
        
    Case 5
        MiDocView1.ActionState = miASTATE_SELECT
        
    Case 7
        On Error GoTo ErrorHand
            MiDocView1.TextSelection.CopyToClipboard
            Exit Sub
ErrorHand:
            MsgBox "Antes de Copiar el Texto, Ejecute el OCR", vbExclamation
    Case 8
        On Error GoTo ErrorHandler
             MiDocView1.ImageSelection.CopyToClipboard
             Exit Sub
ErrorHandler:
                MsgBox "Debe seleccionar para copiar la Imagen", vbExclamation
        
        
    Case 10
        doc.Images(0).Rotate 90
    Case 12
        doc.PrintOut
        

            
End Select

End Sub

Private Sub Command1_Click()
doc.Images(0).Rotate 180

End Sub
Private Sub Doc_OnOCRProgress(ByVal Progress As Long, _
                                 Cancel As Boolean)

  ' Cancel if user has clicked the Cancel button.
  If mblnCancel Then
      Cancel = True
  End If
  
  ' Indicate progress on the ProgressBar control
  'pbrOCRProgress.Value = Progress
  
End Sub


Private Sub Command2_Click()
 ' Set a flag indicating that the user wants to cancel.
  mblnCancel = True

End Sub



Public Function sizeC(xold As Double, yold As Double, xNew As Double, yNew As Double) As Boolean
Dim porx As Double
Dim pory As Double

If xNew < xold Then
        porx = ((xNew * 100) / xold)
        MiDocView1.Width = (((MiDocView1.Width * porx) / 100))
        UserControl.Width = (((UserControl.Width * porx) / 100))
Else
        porx = ((xNew * 100) / xold)
        MiDocView1.Width = (((MiDocView1.Width * porx) / 100))
        UserControl.Width = (((UserControl.Width * porx) / 100))
End If
If yNew < yold Then
        pory = ((yNew * 100) / yold)
        MiDocView1.Height = (((MiDocView1.Height * pory) / 100))
        UserControl.Height = (((UserControl.Height * pory) / 100))
Else
        pory = ((yNew * 100) / yold)
        MiDocView1.Height = (((MiDocView1.Height * pory) / 100))
        UserControl.Height = (((UserControl.Height * pory) / 100))
End If
'MiDocView1.Width = ((MiDocView1.Width * porx) / 100)
'MiDocView1.Height = ((MiDocView1.Height * pory) / 100)
'UserControl.Width = ((UserControl.Width * porx) / 100)
'UserControl.Height = ((UserControl.Height * pory) / 100)

 sizeC = True
End Function

Private Sub UserControl_Resize()
On Error GoTo mesError:
MiDocView1.Width = UserControl.Width - 50
MiDocView1.Height = UserControl.Height - 930
StatusBar1.Width = UserControl.Width
Exit Sub
mesError:
End Sub
