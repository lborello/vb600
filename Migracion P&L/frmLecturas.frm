VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLecturas 
   Caption         =   "Lectura"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   17325
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8655
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   7635
      Begin VB.CommandButton Command8 
         Caption         =   "Guardar y continuar"
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   8040
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cargar Video"
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Top             =   2160
         Width           =   1635
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   435
         Left            =   1200
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.TextBox txtDetalle 
         Height          =   375
         Left            =   4380
         TabIndex        =   20
         Top             =   2280
         Width           =   2235
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Crear Cod"
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   675
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   8160
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   315
         Left            =   540
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   3300
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtCajaPYL 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtVideoLugar 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1035
      End
      Begin VB.CommandButton cmdPlay_Pausa 
         Caption         =   "Espera"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2940
         TabIndex        =   8
         Top             =   420
         Width           =   915
      End
      Begin VB.CommandButton cmdCargarVideo 
         Caption         =   "Cargar Video"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtPasoVideo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Text            =   "C:\P&L\Video\"
         Top             =   900
         Width           =   6075
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   6540
         TabIndex        =   5
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox txtCajaBasa 
         Height          =   375
         Left            =   4860
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   6540
         TabIndex        =   3
         Top             =   1320
         Width           =   315
      End
      Begin VB.TextBox txtNombreVideo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   420
         Width           =   2235
      End
      Begin MSFlexGridLib.MSFlexGrid grdRelacion 
         Height          =   3315
         Left            =   180
         TabIndex        =   1
         Top             =   4740
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   7
      End
      Begin MSDataGridLib.DataGrid grdCajas 
         Height          =   1695
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Video"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Caja Basa"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label lblIDPYL 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   1800
         Width           =   2595
      End
      Begin VB.Label Label1 
         Caption         =   "Caja PYL"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lblCajaAsp 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   375
         Left            =   4860
         TabIndex        =   13
         Top             =   1740
         Width           =   1995
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin WMPLibCtl.WindowsMediaPlayer Video 
      Height          =   9255
      Left            =   0
      TabIndex        =   21
      Top             =   -480
      Width           =   1035
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1826
      _cy             =   16325
   End
End
Attribute VB_Name = "frmLecturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conPYL As New ADODB.Connection
Dim conBasa As New ADODB.Connection
Dim ValorAnteVideo  As Long
Dim PasoVIdeo As String

Private Sub cmdCargarVideo_Click()
Video.URL = txtPasoVideo.Text
Video.Controls.play
ValorAnteVideo = 100
Timer1.Enabled = True

cmdPlay_Pausa.Caption = "Pausa"
txtCajaBasa.Text = ""
lblCajaAsp.Caption = ""
txtCajaPYL.Text = ""
lblIDPYL.Caption = ""
TitulosGrdRelacion
End Sub

Private Sub cmdGuardar_Click()
Guardarvideo
Video.Close
txtCajaBasa.Text = ""
lblCajaAsp.Caption = ""
txtCajaPYL.Text = ""
lblIDPYL.Caption = ""
txtVideoLugar.Text = ""
TitulosGrdRelacion
FileCopy PasoVIdeo & txtNombreVideo.Text, PasoVIdeo & "Procesados\" & Mid(txtNombreVideo, 1, Len(txtNombreVideo) - 4) & " CANT " & grdRelacion.Rows & ".MP4"
If MsgBox("Usted quiere Borrar el Archivo", vbYesNo) = vbYes Then
    Kill PasoVIdeo & txtNombreVideo
End If

txtPasoVideo.Text = ""
txtNombreVideo.Text = ""

End Sub

Private Sub cmdPlay_Pausa_Click()
Play_Pausa cmdPlay_Pausa.Caption
txtCajaBasa.Text = ""
lblCajaAsp.Caption = ""
txtCajaPYL.Text = ""
lblIDPYL.Caption = ""
End Sub

Private Sub Command1_Click()


Dim sql As String
Dim rs As New ADODB.Recordset
    
sql = " SELECT     Caja.Id, Empresa.Nombre, Caja.Numero, Caja.Caja_Asp,CONREF, Fecha , Caja.DETALLE "
sql = sql & "  FROM         Caja INNER JOIN"
sql = sql & "                       Empresa ON Caja.IdEmpresa = Empresa.Id"
sql = sql & "  WHERE    (Caja.Numero LIKE N'%" & Trim(txtCajaPYL.Text) & "%')"
sql = sql & "  ORDER BY Empresa.Nombre, Caja.Numero"


rs.CursorLocation = adUseClient
      
       
rs.Open sql, conPYL
If Not rs.EOF Then
    If Not IsNull(rs!CAJA_ASP) Then
        MsgBox "Ya existe caja asignada"
    End If

End If


Set grdCajas.DataSource = rs.DataSource
       grdCajas.Refresh

 

 If rs.RecordCount = 1 Then
    lblIDPYL.Caption = rs!ID
     txtCajaBasa.SetFocus
    Else
    lblIDPYL.Caption = ""
 End If
 grdCajas.Columns(0).Width = 700
 grdCajas.Columns(1).Width = 800
grdCajas.Columns(2).Width = 2000
grdCajas.Columns(3).Width = 2000
grdCajas.Columns(4).Width = 800
grdCajas.Columns(5).Width = 800
grdCajas.Columns(6).Width = 2000
'grdCajas.Columns(7).Width = 500


End Sub

Private Sub Command2_Click()
 Dim sql As String
 txtDetalle.Text = ""
 
 sql = Dir(PasoVIdeo & "*.mp4")
 txtNombreVideo.Text = sql
 If sql <> "" Then
    txtPasoVideo.Text = PasoVIdeo & sql
    cmdCargarVideo_Click
 Else
    MsgBox "No hay mas archivos "
 End If
 
 
End Sub

Private Sub Command3_Click()

Dim sql As String
Dim caja As Long
Dim Digito As Integer
Dim rs As New ADODB.Recordset
Dim rsPYL As New ADODB.Recordset
If txtCajaBasa.Text <> "" Then

        caja = Mid(txtCajaBasa.Text, 1, Len(txtCajaBasa.Text) - 1)
        Digito = Mid(txtCajaBasa.Text, Len(txtCajaBasa.Text), 1)
        
        
        sql = " SELECT     ID_CAJA, DIGITO_VERIFICADOR, ETIQUETA "
        sql = sql & " From CAJAS "
        sql = sql & " Where ID_CAJA = " & caja
        sql = sql & " And DIGITO_VERIFICADOR = " & Digito
        
        lblCajaAsp.Caption = ""
        rs.Open sql, conBasa
        
                If rs.EOF Then
                   MsgBox "Error en la caja BASE BASA ", vbCritical
                Else
                   
                       sql = " SELECT Caja_Asp "
                       sql = sql & " From Caja"
                       sql = sql & " Where CAJA_ASP = " & rs!ETIQUETA
                       rsPYL.Open sql, conPYL
                       
                       If rsPYL.EOF Then
                           lblCajaAsp.Caption = rs!ETIQUETA
                       Else
                           MsgBox " Ya esta cargada la caja", vbCritical
                       End If
                
                
                End If
        
        Else

 MsgBox "Error en cajas"
End If


End Sub

Private Sub Command4_Click()
Dim Item As String


If Trim(lblIDPYL.Caption) <> "" And Trim(lblCajaAsp.Caption) <> "" Then

Item = "" & vbTab & grdRelacion.Rows & vbTab & lblIDPYL & vbTab & lblCajaAsp & vbTab & txtNombreVideo.Text & vbTab & txtVideoLugar.Text & vbTab & txtDetalle.Text
grdRelacion.AddItem Item

txtCajaBasa.Text = ""
lblCajaAsp.Caption = ""
txtCajaPYL.Text = ""
lblIDPYL.Caption = ""
txtCajaPYL.SetFocus
Else
    MsgBox "eRROR"

End If
txtDetalle.Text = ""

grdRelacion.ScrollTrack = True

End Sub

Private Sub Command5_Click()
Dim sql As String

sql = "INSERT INTO Caja"
sql = sql & " (IdEmpresa, Numero)"
sql = sql & "  VALUES     (39,"
sql = sql & " '" & UCase(Trim(txtCajaPYL.Text)) & "')"

conPYL.Execute sql
MsgBox "Terminado"

End Sub

Private Sub Command6_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim conpyl1 As New ADODB.Connection

conpyl1.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"



sql = " SELECT     TIEMPO_NUMERO,TIEMPO, Id"
sql = sql & " From Caja"
sql = sql & " Where (Not (TIEMPO Is Null))"
rs.Open sql, "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"


Do While Not rs.EOF
    sql = " Update Caja"
    sql = sql & " SET              TIEMPO_NUMERO =" & Replace(Format(rs!TIEMPO, "0000.00"), ",", ".")
    sql = sql & " Where (Not (TIEMPO Is Null)) And ID = " & rs!ID
    conpyl1.Execute sql
    rs.MoveNext
Loop


End Sub

Private Sub Command7_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim Item  As String
sql = " SELECT     Id, Caja_Asp, VIDEO, TIEMPO_NUMERO, DETALLE, Numero"
sql = sql & " From Caja"
sql = sql & "  WHERE VIDEO = '" & Trim(Replace(txtPasoVideo.Text, "C:\P&L\Video\", "")) & "'"
sql = sql & "  ORDER BY VIDEO, TIEMPO_NUMERO"
rs.Open sql, conPYL
TitulosGrdRelacion
cmdPlay_Pausa_Click

Do While Not rs.EOF
    Item = "" & vbTab & grdRelacion.Rows & vbTab & rs!ID & vbTab & rs!CAJA_ASP & vbTab & rs!Video & vbTab & rs!TIEMPO_NUMERO & vbTab & Trim(rs!DETALLE) & vbTab & Trim(rs!numero)
    grdRelacion.AddItem Item
    
    txtVideoLugar.Text = Replace(rs!TIEMPO_NUMERO, ".", ",")
   
    rs.MoveNext
Loop

txtVideoLugar.SetFocus


End Sub

Private Sub Command8_Click()
Guardarvideo
MsgBox "guardaro"
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
    cmdPlay_Pausa_Click
 
 End If
 
End Sub

Private Sub Form_Load()
    Set conPYL = New ADODB.Connection
    conPYL.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    
    Set conBasa = New ADODB.Connection
    conBasa.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
Rem PasoVIdeo = "\\222.15.19.251\f\PLANTA2\Lectura_\P&L\Video\"
PasoVIdeo = "C:\P&L\Video\"
TitulosGrdRelacion


End Sub

Private Sub Form_Resize()

Frame1.Top = 100
Video.Top = 50
Video.Width = frmLecturas.Width - Frame1.Width - 100
Video.Height = frmLecturas.Height - 700
Frame1.Left = Video.Width
End Sub

Private Sub grdCajas_DblClick()
grdCajas.Col = 0
lblIDPYL.Caption = grdCajas.Text

End Sub

Private Sub lblCajaAsp_Change()
If lblCajaAsp.Caption <> "" And lblCajaAsp.Caption <> "" Then
    Command4_Click
End If


End Sub

Private Sub Timer1_Timer()
 txtVideoLugar.Text = Video.Controls.currentPosition
    If ValorAnteVideo = txtVideoLugar.Text Then
        Timer1.Enabled = False
        cmdPlay_Pausa.Caption = "Espera"
    Else
        ValorAnteVideo = txtVideoLugar.Text
    End If
End Sub

Private Sub txtCajaBasa_GotFocus()
lblCajaAsp.Caption = ""
End Sub

Private Sub txtCajaBasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command3_Click
End If
End Sub

Private Sub txtCajaPYL_GotFocus()
lblIDPYL.Caption = ""

End Sub

Private Sub txtCajaPYL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If

End Sub

Private Sub txtVideoLugar_GotFocus()
Timer1.Enabled = False
End Sub

Private Sub txtVideoLugar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Video.Controls.currentPosition = txtVideoLugar.Text
         Play_Pausa "Play"
        Timer1.Enabled = True
            Play_Pausa "Pausa"
    End If

End Sub

Public Sub Play_Pausa(Valor As String)
If Valor = "Play" Then
    Video.Controls.play
    Timer1.Enabled = True
    ValorAnteVideo = 0
    cmdPlay_Pausa.Caption = "Pausa"
    Exit Sub
End If
If Valor = "Pausa" Then
    Video.Controls.pause
    Timer1.Enabled = False
    ValorAnteVideo = 0
    cmdPlay_Pausa.Caption = "Play"
End If
End Sub

Public Sub TitulosGrdRelacion()


    With grdRelacion
        .Clear
        .Rows = 1
        .Cols = 8
        .ColWidth(0) = 100
        .ColWidth(1) = 800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 2000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColAlignment(0) = flexAlignCenterCenter ' flexAlignLeftCenter
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
.ColAlignment(7) = 1
        
        .TextMatrix(0, 1) = "ORDEN"
        .TextMatrix(0, 2) = "ID_CAJA_PYL"
        .TextMatrix(0, 3) = "CAJA_ASP"
        .TextMatrix(0, 4) = "ARCHIVO"
        .TextMatrix(0, 5) = "TIEMPO"
        .TextMatrix(0, 6) = "DETALLE"
    .TextMatrix(0, 7) = "CAJA_PYL"
    End With



End Sub


Public Sub Guardarvideo()
Dim i  As Integer
Dim IDPLY As Long
Dim CAJA_ASP As Double
Dim VideoSTR As String
Dim TIEMPO As String
Dim DETALLE As String
Dim sql As String
Dim TIEMPO_NUMERO As String

For i = 1 To grdRelacion.Rows - 1
 
    IDPLY = grdRelacion.TextMatrix(i, 2)
    CAJA_ASP = grdRelacion.TextMatrix(i, 3)
    VideoSTR = "'" & grdRelacion.TextMatrix(i, 4) & "'"
    TIEMPO = Mid(grdRelacion.TextMatrix(i, 5), 1, 9)
    TIEMPO_NUMERO = Replace(Format(TIEMPO, "0000.00"), ",", ".")
    If Trim(grdRelacion.TextMatrix(i, 6)) <> "" Then
        DETALLE = "'" & Trim(grdRelacion.TextMatrix(i, 6)) & "'"
    Else
        DETALLE = "Null"
    End If
    
    sql = " UPDATE  Caja"
    sql = sql & " SET"
    sql = sql & " Caja_Asp =" & CAJA_ASP
    sql = sql & ", VIDEO =" & VideoSTR
    sql = sql & ", TIEMPO_NUMERO =" & TIEMPO_NUMERO
    sql = sql & ", DETALLE =" & DETALLE
    sql = sql & " Where "
    sql = sql & " ID = " & IDPLY
    conPYL.Execute sql
        
Next
End Sub
