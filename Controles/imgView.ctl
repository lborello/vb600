VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ED512BE6-6629-4FB4-953D-D0C353847163}#1.0#0"; "ImagXpr7.dll"
Begin VB.UserControl ViewImg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   7980
   Begin ImagXpr7Ctl.ImagXpress MiDocView1 
      Height          =   3015
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   5318
      ErrStr          =   "2917BAFF9E86EAF1B61D37759B64E559"
      ErrCode         =   403272359
      ErrInfo         =   1799948618
      Persistence     =   -1  'True
      _cx             =   13467
      _cy             =   5318
      AutoSize        =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SaveTransparencyColor=   0
      OLEDropMode     =   0
      SaveTIFFCompression=   0
      SaveTransparent =   0
      SaveJPEGProgressive=   0   'False
      SaveJPEGGrayscale=   0   'False
      SaveJPEGLumFactor=   10
      SaveJPEGChromFactor=   10
      SaveJPEGSubSampling=   2
      ViewAntialias   =   -1  'True
      BorderType      =   1
      ViewDithered    =   -1  'True
      AlignH          =   1
      AlignV          =   1
      LoadRotated     =   0
      JPEGEnhDecomp   =   -1  'True
      WMFConvert      =   0   'False
      ProcessImageID  =   1
      OwnDIB          =   -1  'True
      FileTimeout     =   10000
      AsyncPriority   =   0
      LZWPassword     =   ""
      ViewUpdate      =   -1  'True
      TWAINProductName=   ""
      TWAINProductFamily=   ""
      TWAINManufacturer=   ""
      TWAINVersionInfo=   ""
      ViewProgressive =   0   'False
      SaveTIFFByteOrder=   0
      FTPUserName     =   ""
      FTPPassword     =   ""
      ProxyServer     =   ""
      SaveEXIFThumbnailSize=   0
      SaveLJPPrediction=   1
      PDFSwapBlackandWhite=   0   'False
      SaveTIFFRowsPerStrip=   0
      TIFFIFDOffset   =   0
      ViewGrayMode    =   0
      SaveWSQQuant    =   1
      DisplayError    =   0   'False
      EvalMode        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3735
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5998
            MinWidth        =   5998
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":1D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":20A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":99F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":A2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":ABAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":B006
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":112A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":116F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":11C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":120DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":12240
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":1255A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":12F6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":13306
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "imgView.ctx":13D96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barImg 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ORC"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desplazar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom+"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zom-"
            Object.ToolTipText     =   "Zom -"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Seleccionar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Texto"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Imagen"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Rotar"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   17
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin VB.ComboBox cboZom 
         Height          =   315
         ItemData        =   "imgView.ctx":14012
         Left            =   6060
         List            =   "imgView.ctx":1402E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Width           =   915
      End
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
               Picture         =   "imgView.ctx":14056
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":144A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":19CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":1A11C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":203B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":278B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "imgView.ctx":28192
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "ViewImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Doc As New MODI.Document
Dim img As New MODI.Image
Dim aImg()  As String

Dim imgActual As Integer
Dim cantImag As Integer

Private mblnCancel As Boolean
Dim Xx As Double
Dim yy As Double

Public Sub MostrarImagen(sImagen As String)
   On Error GoTo mensjError
    Doc.Create (sImagen)
    
    MiDocView1.Document = Doc
    StatusBar1.Panels(1) = sImagen
    StatusBar1.Panels(2) = "Paginas " & imgActual + 1
    StatusBar1.Panels(3) = "Cant " & cantImag + 1
    ID_Archivo
    Exit Sub
mensjError:
    Doc.Create ("")
    MiDocView1.Document = Doc
    StatusBar1.Panels(1) = ""
    StatusBar1.Panels(2) = "Paginas " & 0
    StatusBar1.Panels(3) = "Cant " & cantImag
  
End Sub

Private Sub barImg_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo SALIR
Select Case Button.Index
    Case 1
        On Error GoTo ErrorOCR
        
        Screen.MousePointer = vbHourglass
        Doc.OCR (miLANG_SPANISH)
        Screen.MousePointer = vbDefault
        Exit Sub
        
ErrorOCR:
         MsgBox "No existe Imagen, Verifique", vbCritical, "OCR"
         Screen.MousePointer = vbDefault
    Case 3
        MiDocView1.ActionState = miASTATE_PAN
        
        barImg.Buttons.Item(5).Enabled = True
        
    Case 4
        MiDocView1.ActionState = miASTATE_ZOOM
        Xx = Xx + 0.01
        yy = yy + 0.01
        MiDocView1.SetScale Xx, yy
     Case 5
        MiDocView1.ActionState = miASTATE_ZOOM
        Xx = Xx - 0.01
        yy = yy - 0.01
        MiDocView1.SetScale Xx, yy
        
    Case 6
        MiDocView1.ActionState = miASTATE_SELECT
        
    Case 8
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
        Doc.Images(0).Rotate 90
    Case 13
     Dim CON  As New ADODB.Connection
        CON.Open strConBasa
    If MsgBox("Usted quiere marcarlo como Impreso", vbYesNo) = vbYes Then
        CON.Execute " Update DOCUMENTOS_DIGITALES Set IMPRESO = '1' Where ID =" & StatusBar1.Panels.Item(4).Text
    Else
        CON.Execute " Update DOCUMENTOS_DIGITALES Set IMPRESO = '0' Where ID =" & StatusBar1.Panels.Item(4).Text
    End If
    
        Doc.PrintOut
        
    Case 14
            If imgActual = 0 Then
                Me.MostrarImagen (aImg(0))
            Else
                imgActual = imgActual - 1
                Me.MostrarImagen (aImg(imgActual))
                
            End If
    Case 15
            If imgActual = 0 Then
                Me.MostrarImagen (aImg(0))
            Else
                imgActual = imgActual - 1
                Me.MostrarImagen (aImg(imgActual))
                
            End If
    
    
    
    Case 16
            If imgActual = cantImag Then
                Me.MostrarImagen (aImg(cantImag))
            Else
                imgActual = imgActual + 1
                Me.MostrarImagen (aImg(imgActual))
                
            End If
        
Case 17
    Dim NombreImpresora As String
           InputBox "Ingrese el Nombre de la impresora", "Impresora", "ExportPDF"
            Doc.PrintOut , , , NombreImpresora
 Case 18
    Dim Paso As String
    Clipboard.Clear
           Clipboard.SetText CStr(StatusBar1.Panels.Item(1).Text)
           MsgBox "El paso fue copiado"
           
End Select
SALIR:
End Sub

Private Sub Command1_Click()
    Doc.Images(0).Rotate 180

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
        MiDocView1.width = (((MiDocView1.width * porx) / 100))
        UserControl.width = (((UserControl.width * porx) / 100))
        StatusBar1.Panels(1).width = ((StatusBar1.Panels(1).width * porx) / 100)
        StatusBar1.Panels(2).width = ((StatusBar1.Panels(2).width * porx) / 100)
        StatusBar1.Panels(3).width = ((StatusBar1.Panels(3).width * porx) / 100)
Else
        porx = ((xNew * 100) / xold)
        MiDocView1.width = (((MiDocView1.width * porx) / 100))
        UserControl.width = (((UserControl.width * porx) / 100))
        StatusBar1.Panels(1).width = ((StatusBar1.Panels(1).width * porx) / 100)
        StatusBar1.Panels(2).width = ((StatusBar1.Panels(2).width * porx) / 100)
        StatusBar1.Panels(3).width = ((StatusBar1.Panels(3).width * porx) / 100)
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

 sizeC = True
End Function



Private Sub cboZom_Click()
Select Case cboZom.Text
Case 10
        Xx = 0.1
        yy = 0.1
Case 30
        Xx = 0.3
        yy = 0.3
Case 60
        Xx = 0.6
        yy = 0.6
Case 90
        Xx = 0.9
        yy = 0.9
Case 100
        Xx = 1
        yy = 1
Case 150
        Xx = 1.5
        yy = 1.5
Case 200
        Xx = 2
        yy = 2
Case 300
        Xx = 3
        yy = 3
End Select
        MiDocView1.SetScale Xx, yy
End Sub

Private Sub UserControl_Initialize()
Xx = 0.77
 yy = 0.77
     Inicio
End Sub

Private Sub UserControl_Resize()
On Error GoTo mesError:
Debug.Print UserControl.width
Debug.Print UserControl.Height
    MiDocView1.width = UserControl.width - 50
    MiDocView1.Height = UserControl.Height - 930
    StatusBar1.Panels(2).width = 900
    StatusBar1.Panels(3).width = 900
    StatusBar1.Panels(1).width = UserControl.width - 2500
    UserControl.Refresh
Exit Sub
mesError:
End Sub
Public Sub MostrarImagenes(a() As String)
    aImg = a
    cantImag = UBound(aImg)
    Me.MostrarImagen (aImg(0))
    ID_Archivo
End Sub


Public Sub ID_Archivo()
 Dim i As Integer
 Dim f As Integer
 f = Len(StatusBar1.Panels.Item(1).Text)
 For i = 0 To Len(StatusBar1.Panels.Item(1).Text)
    
    If Mid(StatusBar1.Panels.Item(1).Text, f, 1) = "\" Then
        Exit For
        
    End If
    Debug.Print Mid(StatusBar1.Panels.Item(1).Text, f, 1)
    f = f - 1
 
 Next
 
 Dim dato As String
 
    dato = Mid(StatusBar1.Panels.Item(1).Text, f + 1)
    dato = Replace(dato, ".tif", "")
 dato = Replace(dato, ".TIF", "")
 StatusBar1.Panels.Item(4).Text = dato
 
End Sub
