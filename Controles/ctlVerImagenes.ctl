VERSION 5.00
Object = "{514ACB2B-3B06-4ABB-8ECE-D18141FC9562}#1.0#0"; "SSICRXpr4.dll"
Object = "{ED512BE6-6629-4FB4-953D-D0C353847163}#1.0#0"; "ImagXpr7.dll"
Begin VB.UserControl ctlVerImagenes 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13500
   ScaleHeight     =   4545
   ScaleWidth      =   13500
   Begin VB.TextBox txtPaginaFija 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4980
      TabIndex        =   9
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdAdelante 
      Caption         =   ">"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   60
      Width           =   255
   End
   Begin VB.CommandButton cmdAtras 
      Caption         =   "<"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   60
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1035
   End
   Begin ImagXpr7Ctl.ImagXpress ixpControlImagen 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Controles con Doble click"
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6271
      ErrStr          =   "2917BAFF9E86EAF1B61D37759B64E559"
      ErrCode         =   263405616
      ErrInfo         =   2021759087
      Persistence     =   -1  'True
      _cx             =   14843
      _cy             =   6271
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
      ScrollBars      =   3
      AlignH          =   0
      AlignV          =   0
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
   Begin VB.Label lblPaso 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      Height          =   315
      Left            =   4260
      TabIndex        =   8
      Top             =   0
      Width           =   5715
   End
   Begin VB.Label lblPaginas 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6420
      TabIndex        =   5
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "de"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   60
      Width           =   315
   End
   Begin VB.Label lblPagina 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5580
      TabIndex        =   3
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Pagina:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3660
      TabIndex        =   2
      Top             =   60
      Width           =   1155
   End
   Begin SSXICR4LibCtl.SSICRXpr SSICRXpr1 
      Left            =   360
      Top             =   720
      _cx             =   635
      _cy             =   635
      Enabled         =   -1  'True
      ClassifierPath  =   ""
      DisplayError    =   0   'False
      ImageSource     =   0
      InkColor        =   0
      OwnDIB          =   -1  'True
      EvalMode        =   0
      RaiseExceptions =   0   'False
   End
End
Attribute VB_Name = "ctlVerImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Zoom As Double
Dim I_ScrollX As Double
Dim I_ScrollY As Double

Private Sub cmdAdelante_Click()
    If ixpControlImagen.Pages = ixpControlImagen.PageNbr Then
        ixpControlImagen.PageNbr = ixpControlImagen.Pages
    Else
        ixpControlImagen.PageNbr = ixpControlImagen.PageNbr + 1
    End If
    PonerImagen ixpControlImagen.FileName, ixpControlImagen.PageNbr
    ixpControlImagen.Refresh
    lblPagina.Caption = ixpControlImagen.PageNbr
    lblPaginas.Caption = ixpControlImagen.Pages
End Sub

Private Sub cmdAtras_Click()
    If ixpControlImagen.PageNbr > 1 Then
        ixpControlImagen.PageNbr = ixpControlImagen.PageNbr - 1
    End If
    PonerImagen ixpControlImagen.FileName, ixpControlImagen.PageNbr
    ixpControlImagen.Refresh
    lblPagina.Caption = ixpControlImagen.PageNbr
    lblPaginas.Caption = ixpControlImagen.Pages

End Sub

'Private Sub Command1_Click()
'ixpControlImagen.Rotate 180
'
'
'
'Dim i As ImagXpress
'Set i = ixpControlImagen
'
'i.SaveFileName = "C:\TempImagen\" & lblPagina & ".tif"
'i.SaveMultiPage = False
'i.SaveFileType = FT_TIFF
'i.SaveTIFFByteOrder = TIFF_INTEL
'Dim d As New MODI.Document
'Dim im As New MODI.Image
'
'
'd.Create
'
'Set im.Picture = ixpControlImagen.Picture
'd.Images.Add im, im
'd.SaveAs "c:\LUIS.TIF"
'
'
'
'
'
'Rem ixpControlImagen.SaveFileName = ixpControlImagen.FileName
'
'End Sub

Private Sub ixpControlImagen_DblClick()
    ixpControlImagen.ToolbarActivated = True
End Sub

Private Sub ixpControlImagen_Scroll(ByVal Bar As Integer, ByVal Action As Integer)
    I_ScrollX = ixpControlImagen.ScrollX
    I_ScrollY = ixpControlImagen.ScrollY

End Sub

Private Sub tbrControlImagen_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Caption
    Case "Zom+"
        Zoom = Zoom + 0.05
        ixpControlImagen.Zoom Zoom
        ixpControlImagen.ScrollX = I_ScrollX
        ixpControlImagen.ScrollY = I_ScrollY
    Case "Zom-"
        Zoom = Zoom - 0.05
        ixpControlImagen.Zoom Zoom
       
            ixpControlImagen.ScrollX = I_ScrollX
        ixpControlImagen.ScrollY = I_ScrollY
    End Select
End Sub

Public Function PonerImagen(Paso As String, Optional Pagina As Integer)
ixpControlImagen.FileName = Paso
lblPaso.Caption = Paso
If txtPaginaFija.Text <> "" Then
Pagina = txtPaginaFija.Text
Else
Pagina = 1
End If

ixpControlImagen.PageNbr = Pagina
lblPagina.Caption = ixpControlImagen.PageNbr
lblPaginas.Caption = ixpControlImagen.Pages

 Rem ixpControlImagen.ScrollX = I_ScrollX
 
 Dim i As Integer
' If chkScroolSuperior.Value = 1 Then
'    ixpControlImagen.ScrollX = 0
'    ixpControlImagen.ScrollY = 0
'Else
'      If chkScroolInferior.Value = 1 Then
'            ixpControlImagen.ScrollX = 0
'            Rem Debug.Print ixpControlImagen.ScrollY
'           Rem  ixpControlImagen.AlignV = ALIGNV_Bottom
'
'
'        Else
'
'            Debug.Print ixpControlImagen.ScrollBarLargeChangeH
'            Debug.Print ixpControlImagen.ScrollBarLargeChangeV
'            Debug.Print ixpControlImagen.ScrollBarSmallChangeH
'            Debug.Print ixpControlImagen.ScrollBarSmallChangeV
'            Debug.Print ixpControlImagen.ScrollBarSmallChangeV
'        End If
'
'End If

ixpControlImagen.ScrollX = I_ScrollX
            ixpControlImagen.ScrollY = I_ScrollY
            
    ixpControlImagen.AlignH = ALIGNH_Center
    ixpControlImagen.Refresh
    ixpControlImagen.ScrollBars = SB_None
    ixpControlImagen.ScrollBars = SB_Both

 
  

End Function

Public Function Resise(Height As Long, width As Long)
    UserControl.width = width
    UserControl.Height = Height
End Function

Private Sub UserControl_Initialize()
  Inicio
End Sub

Private Sub UserControl_Resize()
On Error GoTo salir
    ixpControlImagen.width = UserControl.width - 100
    ixpControlImagen.Height = UserControl.Height - 300
salir:
End Sub
Public Function ZonnFijo(Zoom As Double)
  ixpControlImagen.Zoom Zoom
End Function


Public Sub Save_Imagen(Paso As String)
Dim i As Integer
For i = 1 To ixpControlImagen.Pages

PonerImagen ixpControlImagen.FileName, i
 If i = 1 Then
    ixpControlImagen.SaveMultiPage = False
 Else
    ixpControlImagen.SaveMultiPage = True
 End If
    ixpControlImagen.SaveFileName = Paso
    ixpControlImagen.SaveFileType = FT_TIFF_G4
    


ixpControlImagen.SaveFile


Next

End Sub
