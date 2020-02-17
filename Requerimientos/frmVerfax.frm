VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVerfax 
   Caption         =   "Ver Fax"
   ClientHeight    =   8520
   ClientLeft      =   135
   ClientTop       =   270
   ClientWidth     =   12615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   12615
   Begin VB.PictureBox ImgThumbnail1 
      BackColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   3180
      ScaleHeight     =   855
      ScaleWidth      =   3255
      TabIndex        =   34
      Top             =   7080
      Width           =   3315
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   435
      Left            =   9540
      TabIndex        =   33
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5115
      Left            =   9060
      TabIndex        =   14
      Top             =   900
      Width           =   3675
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   435
         Left            =   2400
         TabIndex        =   32
         Top             =   4500
         Width           =   1035
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Siguiente"
         Height          =   495
         Left            =   1260
         TabIndex        =   31
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdClienteNocorres 
         Caption         =   "Cliente no Corre."
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtPaso 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imagen"
         Height          =   315
         Left            =   1020
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Paso"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fijar Valor - F12"
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txtRemito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Saltar Imagen"
         Height          =   495
         Left            =   2100
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtRemito2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   1380
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Dos Remitos"
         Height          =   315
         Left            =   2100
         TabIndex        =   20
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Recargar"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   2460
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Firma ilegible"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "No esta en la lista"
         Height          =   495
         Left            =   1260
         TabIndex        =   17
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "existe"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   3900
         Width           =   975
      End
      Begin VB.CheckBox chkNoentrodoLista 
         Caption         =   "No encontrado"
         Height          =   555
         Left            =   1140
         TabIndex        =   16
         Top             =   2280
         Width           =   1395
      End
      Begin Requerimientos.ctlClientes ctlClientes1 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
      End
      Begin Requerimientos.ctlClienteUsuario ctlClienteUsuario1 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   2880
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   873
      End
      Begin VB.Label lblRemito 
         Caption         =   "Label2"
         Height          =   315
         Left            =   2700
         TabIndex        =   29
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label lblSector 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   435
         Left            =   120
         TabIndex        =   28
         Top             =   3360
         Width           =   3255
      End
   End
   Begin VB.PictureBox ImgEdit1 
      Height          =   2355
      Left            =   8880
      ScaleHeight     =   2295
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   495
      Left            =   11340
      TabIndex        =   13
      Top             =   7860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   435
      Left            =   11340
      TabIndex        =   12
      Top             =   7260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   555
      Left            =   11220
      TabIndex        =   11
      Top             =   6420
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   495
      Left            =   11220
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   435
      Left            =   11100
      TabIndex        =   9
      Top             =   5460
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12120
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox oleImgEdit1 
      Height          =   6030
      Left            =   480
      ScaleHeight     =   5970
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":0634
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":0A86
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":0ED8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":132A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":177C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":1A96
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":1DB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":2202
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":2654
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":296E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":2DC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":3212
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":3664
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":397E
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":3DD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVerfax.frx":4222
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox oleImgAdmin1 
         Height          =   480
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   35
         Top             =   0
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1429
      ButtonWidth     =   1561
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ver +"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ver -"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rotar D"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rotar I"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vertical"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sello"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ant Pagina"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sig Pagina"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fijar- F12"
         EndProperty
      EndProperty
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   8520
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   315
         Left            =   6780
         TabIndex        =   5
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   375
         Left            =   8820
         TabIndex        =   4
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.PictureBox oleImgThumbnail1 
      BackColor       =   &H8000000A&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   2895
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2955
   End
End
Attribute VB_Name = "frmVerfax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Selection As Boolean 'Selection = True, selection rect drawn.
Dim Annot8Visible As Boolean 'Annot8Visible = True, annotation toolbox is
                            'visible
Dim CurrentPage As Integer 'CurPage = currently displayed image page
Dim LastPage As Integer 'LastPage = last page viewed before current page
Dim TotalPages As Integer 'TotalPages = image document page count
Dim numbits As Integer 'number of bits per pixel supported by this device
Dim PasoImagen As String
Dim DataTimeFax As Date

'Const defines
Const NoTool = 0
Const AnnoSelection = 1
Const AnnoFreehand = 2
Const AnnoHiLight = 3
Const AnnoStraightLine = 4
Const AnnoHollowRect = 5
Const AnnoFilledRect = 6
Const AnnoText = 7
Const AnnoAttachNote = 8
Const AnnoTextFromFile = 9
Const AnnoRubberStamp = 10
Const BestFit = 0
Const FitWidth = 1
Const FitHeight = 2
Const InchToInch = 3
Const ErrCancel = 32755
Const ZoomMax = 6554
Const ZoomMin = 2
Const TiffImage = 1
Const AwdImage = 2
Const BmpImage = 3
Const ImageChanged = "Image has changed.  Do you want to save changes?"
Dim PositionX As Long
Dim PositionY As Long
Dim ZoomRemito As Long
Dim PasoTotal As String
Public Paso As String
Dim Left1, Top1, Width1, Height1 As Long
  Dim rsRemitos As ADODB.Recordset

Private Sub Imprimir()
  
  On Error Resume Next 'handle errors ourselves
  oleImgEdit1.Zoom = 100
  oleImgEdit1.PrintImage oleImgEdit1.Page, oleImgEdit1.Page
  If oleImgEdit1.StatusCode <> 0 Then
    MsgBox Err.Description + " Code = " + Hex(oleImgEdit1.StatusCode), 16
  End If

End Sub


Public Function PonerImagen(Imagen As String) As Boolean
  Rem MoverImagen = True
  Dim temp As String 'image name and path
        
  On Error GoTo LUIS 'handle errors ourselves incase of cancel
 Rem  oleImgAdmin1.Flags = 0 'clear Flags
  If oleImgEdit1.ImageModified = True Then
'    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
'      Rem mnuSave_Click
''      If Err = ErrCancel Then '32755 = Cancel pressed
''        Exit Function
''      End If
'    End If
  End If
 Rem  oleImgAdmin1.Filter = "All Image Files|*.tif;*.bmp;*.jpg;*.pcx;*.dcx|TIFF files (*.tif)|*.tif|BMP files (*.bmp)|*.bmp|PCX/DCX Document (*.pcx, *.dcx)|*.pcx;*.dcx|JPG File (*.jpg)|*.jpg|All Files (*.*)|*.*|"
  Rem oleImgAdmin1.ShowFileDialog 0, frmSample.hWnd
 Rem If Err = ErrCancel Then '32755 = Cancel pressed
 Rem    Exit Function
 Rem  End If
' Rem  If oleImgAdmin1.StatusCode <> 0 Then
'    MsgBox Err.Description + " Code = " + Hex(oleImgAdmin1.StatusCode), 16
'    Exit Function
'  End If

  
    temp = Trim(Imagen)
  oleImgEdit1.Zoom = 38
Noimgen:
  oleImgEdit1.Image = temp
  oleImgThumbnail1.Image = oleImgEdit1.Image
  If numbits > 8 Then 'video driver supports hicolor or truecolor
    oleImgEdit1.ImagePalette = 3 'Set for 24 bit RGB.
  End If
  oleImgEdit1.ImagePalette = wiPaletteGray8
  oleImgEdit1.Page = 1
 Rem oleImgEdit1.Display
  TotalPages = oleImgEdit1.PageCount
  oleImgThumbnail1.ThumbSelected(1) = True
        
  'Now that we have an image, enable the needed menus.
 
  Exit Function
LUIS:
  If Err.Number = 53 Then
    
    temp = "C:\noimagen.tif"
    GoTo Noimgen
   Rem  MoverImagen = False
  End If

End Function

Private Sub RemitoImagen()
    Dim MyName As String
    MyName = Dir(txtPaso & "*.Tif", vbDirectory)
    If MyName <> "" Then
       PonerImagen txtPaso & MyName
       PasoTotal = txtPaso & MyName
    Else
    PasoTotal = ""
    End If
    oleImgEdit1.ScrollPositionX = PositionX
    oleImgEdit1.ScrollPositionY = PositionY
    If ZoomRemito <> 0 Then
        oleImgEdit1.Zoom = ZoomRemito
    End If
    oleImgEdit1.Refresh
   
End Sub

Private Sub cmdClienteNocorres_Click()
Dim SQL As String
SQL = " Update REMITOS_CUERPO Set COD_USUARIO_CLIENTE = 10000"
SQL = SQL & " Where NRO_remito = " & rsRemitos!NRO_REMITO
ConBasa.Execute SQL
PonerimagenRemito
End Sub

Private Sub Command1_Click()
RemitoImagen
End Sub

Private Sub Command10_Click()
Dim SQL As String
SQL = " Update REMITOS_CUERPO Set COD_USUARIO_CLIENTE = " & ctlClienteUsuario1.Valor
SQL = SQL & " Where NRO_remito = " & rsRemitos!NRO_REMITO
ConBasa.Execute SQL
ctlClienteUsuario1.Clear
PonerimagenRemito
ctlClienteUsuario1.SetFocus
End Sub

Private Sub Command11_Click()
    Set rsRemitos = New ADODB.Recordset
    Dim SQL As String
    Dim existe As String
        SQL = "SELECT COD_USUARIO_CLIENTE, IMAGEN, NRO_REMITO"
        SQL = SQL & " From REMITOS_CUERPO order by NRO_REMITO "
        rsRemitos.Open SQL, ConBasa
        Do While Not rsRemitos.EOF
            existe = Dir("\\Server1basa\remitos\Remito " & rsRemitos!NRO_REMITO & ".tif")
            If existe <> "" Then
                FileCopy "\\Server1basa\remitos\Remito " & rsRemitos!NRO_REMITO & ".tif", "\\Server1basa\remitos\ok\Remito " & rsRemitos!NRO_REMITO & ".tif"
                Kill "\\Server1basa\remitos\Remito " & rsRemitos!NRO_REMITO & ".tif"
            End If
            rsRemitos.MoveNext
        Loop
End Sub

Private Sub Command12_Click()
    ImgEdit1.ClipboardPaste
  
End Sub

Private Sub Command13_Click()
  
  
  
  
  
  Rem MoverImagen = True
  Dim temp As String 'image name and path
  temp = Trim("D:\firma.tif")
  On Error GoTo LUIS 'handle errors ourselves incase of cancel
 Rem  oleImgAdmin1.Flags = 0 'clear Flags
  If ImgEdit1.ImageModified = True Then
'    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
'      Rem mnuSave_Click
''      If Err = ErrCancel Then '32755 = Cancel pressed
''        Exit Function
''      End If
'    End If
  End If
 Rem  oleImgAdmin1.Filter = "All Image Files|*.tif;*.bmp;*.jpg;*.pcx;*.dcx|TIFF files (*.tif)|*.tif|BMP files (*.bmp)|*.bmp|PCX/DCX Document (*.pcx, *.dcx)|*.pcx;*.dcx|JPG File (*.jpg)|*.jpg|All Files (*.*)|*.*|"
  Rem oleImgAdmin1.ShowFileDialog 0, frmSample.hWnd
 Rem If Err = ErrCancel Then '32755 = Cancel pressed
 Rem    Exit Function
 Rem  End If
' Rem  If oleImgAdmin1.StatusCode <> 0 Then
'    MsgBox Err.Description + " Code = " + Hex(oleImgAdmin1.StatusCode), 16
'    Exit Function
'  End If

  
    Rem temp = Trim(Imagen)
  ImgEdit1.Zoom = 38
Noimgen:
  ImgEdit1.Image = temp
  oleImgThumbnail1.Image = ImgEdit1.Image
  If numbits > 8 Then 'video driver supports hicolor or truecolor
    ImgEdit1.ImagePalette = 3 'Set for 24 bit RGB.
  End If
  ImgEdit1.Page = 1
  ImgEdit1.Display
  TotalPages = ImgEdit1.PageCount
  oleImgThumbnail1.ThumbSelected(1) = True
 'Now that we have an image, enable the needed menus.
   Exit Sub
LUIS:
  If Err.Number = 53 Then
    temp = "C:\noimagen.tif"
    GoTo Noimgen
    Rem  MoverImagen = False
  End If

  
End Sub

Private Sub Command14_Click()
ImgEdit1.Page = 2
ImgEdit1.Display
End Sub

Private Sub Command15_Click()
     Dim Con As New ADODB.Connection
        
       Dim SQL As String
        Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\montemar\montemar.mdb;Persist Security Info=False"
       
    
    Dim MiNombre As String
    
MiNombre = Dir(txtPaso, vbDirectory)   ' Recupera la primera entrada.
Do While MiNombre <> ""   ' Inicia el bucle.
   ' Ignora el directorio actual y el que lo abarca.
   If MiNombre <> "." And MiNombre <> ".." Then
    SQL = "insert into archivo (archivo,Cantidad,Paso) values ('" & MiNombre & "'," & oleImgEdit1.PageCount & ",'" & txtPaso & "')"
    Con.Execute SQL
    PonerImagen txtPaso & MiNombre
    Debug.Print txtPaso & vbTab & MiNombre & vbTab & oleImgEdit1.PageCount
   End If
   MiNombre = Dir   ' Obtiene siguiente entrada.
Loop


    
End Sub

Private Sub Command16_Click()
ImgEdit1.Zoom = ImgEdit1.Zoom + 10
MsgBox ImgEdit1.Zoom
ImgEdit1.Display
End Sub

Private Sub Command17_Click()
ImgEdit1.Save
End Sub

Private Sub Command18_Click()
Set rsRemitos = New ADODB.Recordset
    Dim SQL As String
    
SQL = " SELECT NRO_REMITO, IMAGEN, COD_USUARIO_CLIENTE, TIPO, ID_CLIENTE"
SQL = SQL & " From REMITOS_CUERPO"
SQL = SQL & vbCrLf & "  WHERE  (IMAGEN IS NULL) "
SQL = SQL & vbCrLf & " ORDER BY NRO_REMITO DESC"
rsRemitos.Open SQL, ConBasa
Set ctlClienteUsuario1.Conexion = ConBasa


Do While Not rsRemitos.EOF
Dim existe As String
existe = Dir("\\Server1basa\remitos\Remito " & rsRemitos!NRO_REMITO & ".tif")
If existe <> "" Then
     ConBasa.Execute "Update REMITOS_CUERPO SET IMAGEN = 'NI' Where NRO_REMITO =" & rsRemitos!NRO_REMITO
   
End If
rsRemitos.MoveNext
Loop
End Sub

Private Sub Command2_Click()
    CommonDialog1.ShowOpen
    Dim largo As Integer
    Dim PasoLargo As Integer
    
    largo = Len(CommonDialog1.FileTitle)
    PasoLargo = Len(CommonDialog1.FileName)
    txtPaso.Text = Mid(CommonDialog1.FileName, 1, PasoLargo - largo)
End Sub

Private Sub Command3_Click()
    PositionX = oleImgEdit1.ScrollPositionX
    PositionY = oleImgEdit1.ScrollPositionY
    ZoomRemito = oleImgEdit1.Zoom
End Sub

Private Sub Command4_Click()
   If Dir(PasoTotal) <> "" Then
        CopiarRemito PasoTotal, True, "\\Server1basa\remitos\Para Revisar\"
        txtRemito.Text = ""
        RemitoImagen
        txtRemito.SetFocus
   End If
End Sub

Private Sub Command5_Click()

            Dim SQL As String
            SQL = " UPDATE REMITOS_CUERPO SET IMAGEN = 'SI' Where NRO_REMITO =" & txtRemito.Text
            ConBasa.Execute SQL
            CopiarRemito PasoTotal, False
            txtRemito.Text = ""
            SQL = " UPDATE REMITOS_CUERPO SET IMAGEN = 'SI' Where NRO_REMITO =" & txtRemito2.Text
            ConBasa.Execute SQL
            CopiarRemito PasoTotal, True
            txtRemito2.Text = ""
            RemitoImagen
            txtRemito.SetFocus

End Sub

Private Sub Command6_Click()
    Set rsRemitos = New ADODB.Recordset
    Dim SQL As String
    
SQL = " SELECT NRO_REMITO, IMAGEN, COD_USUARIO_CLIENTE, TIPO, ID_CLIENTE"
SQL = SQL & " From REMITOS_CUERPO"
SQL = SQL & vbCrLf & "  WHERE  (NOT (IMAGEN IS NULL)) AND  (ID_CLIENTE =" & ctlClientes1.Cod_CLiente & " ) "
If chkNoentrodoLista.Value = 1 Then
    SQL = SQL & vbCrLf & "  AND COD_USUARIO_CLIENTE = 1000"
Rem SQL = SQL & vbCrLf & "  AND COD_USUARIO_CLIENTE = 10000"
    Else
   SQL = SQL & vbCrLf & "  AND (COD_USUARIO_CLIENTE is null)"
End If
SQL = SQL & vbCrLf & " ORDER BY NRO_REMITO DESC"
rsRemitos.Open SQL, ConBasa
Set ctlClienteUsuario1.Conexion = ConBasa
ctlClienteUsuario1.Cliente = ctlClientes1.Cod_CLiente
PonerimagenRemito


End Sub

Private Sub Command7_Click()
Dim SQL As String
SQL = " Update REMITOS_CUERPO Set COD_USUARIO_CLIENTE = 0"
SQL = SQL & " Where NRO_remito = " & rsRemitos!NRO_REMITO
ConBasa.Execute SQL
PonerimagenRemito
End Sub

Private Sub Command8_Click()
Dim SQL As String
SQL = " Update REMITOS_CUERPO Set COD_USUARIO_CLIENTE = 1000"
SQL = SQL & " Where NRO_remito = " & rsRemitos!NRO_REMITO
ConBasa.Execute SQL
PonerimagenRemito
End Sub

Private Sub Command9_Click()
    PonerimagenRemito
End Sub

Private Sub ctlClienteUsuario1_SectorEncontrado(Sector As String)
lblSector = Sector
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErroSalir
If KeyCode = 123 Then
     PositionX = oleImgEdit1.ScrollPositionX
    PositionY = oleImgEdit1.ScrollPositionY
    ZoomRemito = oleImgEdit1.Zoom

End If
ErroSalir:

End Sub

Private Sub Form_Load()


Dim rs As ADODB.Recordset
Dim SQL As String
 If (CRequerimientos.Count) = 0 Then
    Set ConBasa = New ADODB.Connection
    UserName = "BASA"
    Password = "1742"
    ConBasa.Open "Provider=MSDAORA.1;Password=1742;User ID=basa;Data Source=bpdc;Persist Security Info=True"
    Exit Sub
 End If
If Paso = "" Then
    SQL = " SELECT"
    SQL = SQL & " REQ.IDRequerimiento , FAX.Path"
    SQL = SQL & " From"
    SQL = SQL & " REQUERIMIENTO REQ,"
    SQL = SQL & " FAX FAX"
    SQL = SQL & " Where"
    SQL = SQL & " ( REQ.IDFAX=FAX.IDFAX )"
    SQL = SQL & " AND  REQ.IDRequerimiento = " & CLng(CRequerimientos.Item(1).NumeroRequerimiento)
    Set rs = New ADODB.Recordset
    rs.Open SQL, ConBasa
    If Not rs.EOF Then
     PonerImagen CStr(rs!Path)
    Else
        MsgBox "Este Requerimiento no tiene imagen"
        Unload Me
    End If
Else
PonerImagen CStr(Paso)
End If
End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim Zoon As Integer
   On Error GoTo SALIR:
    Select Case Button
    Case "Rotar D"
          oleImgEdit1.RotateLeft
    Case "Vertical"
          oleImgEdit1.Flip
    Case "Rotar I"
        oleImgEdit1.RotateRight
    Case "Ver +"
        Zoo = Zoo + 10
        oleImgEdit1.Zoom = Zoo
    Case "Ver -"
        If Zoo < 20 Then
            Zoo = 20
        Else
            Zoo = Zoo - 10
            oleImgEdit1.Zoom = Zoo
        End If
    Case "Imprimir"
        Imprimir
    Case "Sello"
            Dim rs As ADODB.Recordset
            Dim SQL As String
            Dim Responsable As String
            SQL = " SELECT P.Nombre "
            SQL = SQL & " ,P.Apellido  "
            SQL = SQL & " From "
            SQL = SQL & " REQUERIMIENTO REQ,"
            SQL = SQL & " Personal P "
            SQL = SQL & " Where"
            SQL = SQL & " ( REQ.IDpersonal=P.idPersonal )"
            SQL = SQL & " AND  REQ.IDRequerimiento = " & CLng(CRequerimientos.Item(1).NumeroRequerimiento)
            Set rs = New ADODB.Recordset
            rs.Open SQL, ConBasa
            If Not rs.EOF Then
                Responsable = UCase(CStr(rs!Nombre)) & "  " & UCase(CStr(rs!Apellido))
            Else
                Responsable = "No"
            End If
            oleImgEdit1.AnnotationType = wiTextStamp
            oleImgEdit1.AnnotationFont.Name = "ARIAL"
            oleImgEdit1.AnnotationFont.Size = 20
            oleImgEdit1.AnnotationFont.Weight = 300
            oleImgEdit1.AnnotationFontColor = &HFF&
            oleImgEdit1.AnnotationStampText = "Numero de Requerimiento = " & CRequerimientos.Item(1).NumeroRequerimiento & vbCrLf & "Responsable:" & Responsable & vbCrLf & "Fecha Y Hora :" & SysDateCompare
            oleImgEdit1.Refresh
    
        Case "Ant Pagina"
            If oleImgEdit1.Page > 1 Then
                oleImgEdit1.Page = oleImgEdit1.Page - 1
                oleImgEdit1.Display
            End If
        Case "Sig Pagina"
            If oleImgEdit1.PageCount > oleImgEdit1.Page Then
                oleImgEdit1.Page = oleImgEdit1.Page + 1
                oleImgEdit1.Display
            End If
    
    
    End Select
   oleImgEdit1.Refresh
SALIR:
End Sub

Public Sub CopiarRemito(PasoInicio As String, Borrar As Boolean, Optional destino As String)
Dim existe As String
   existe = Dir("\\Server1basa\remitos\Remito " & txtRemito.Text & ".tif")
    If existe = "" Then
        If destino = "" Then
            FileCopy PasoInicio, "\\Server1basa\remitos\Remito " & txtRemito.Text & ".tif"
          Else
              FileCopy PasoInicio, destino & Format(Date, "dd_mm_yyyy") & " " & Format(Time, "hh_nn_ss") & ".tif"
                 
        End If
        
        If Borrar = True Then
            If Dir(PasoInicio) <> "" Then
                Kill PasoInicio
            End If
        
        End If
    Else
        MsgBox "Ya existe"
        If MsgBox("Usted desea borrarlo", vbYesNo) = vbYes Then
           If Dir(PasoInicio) <> "" Then
                Kill PasoInicio
            End If
        End If
    End If
End Sub

Private Sub txtRemito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Dim SQL As String
            SQL = " UPDATE REMITOS_CUERPO SET IMAGEN = 'SI' Where NRO_REMITO =" & txtRemito.Text
            ConBasa.Execute SQL
            CopiarRemito PasoTotal, True
            txtRemito.Text = ""
            RemitoImagen
            txtRemito.SetFocus
    End If

End Sub

Public Sub PonerimagenRemito()
Dim existe As String
On Error GoTo SALIR
rsRemitos.MoveNext
existe = Dir("\\Server1basa\remitos\Remito " & rsRemitos!NRO_REMITO & ".tif")

If existe <> "" Then
       PonerImagen "\\Server1basa\remitos\Remito " & rsRemitos!NRO_REMITO & ".tif"
lblRemito = rsRemitos!NRO_REMITO
    oleImgEdit1.ScrollPositionX = PositionX
    oleImgEdit1.ScrollPositionY = PositionY
    If ZoomRemito <> 0 Then
        oleImgEdit1.Zoom = ZoomRemito
    End If
    oleImgEdit1.Refresh
    
    
    Exit Sub
SALIR:
 MsgBox "No existen registro"
   
End If


End Sub
