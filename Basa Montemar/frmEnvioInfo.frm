VERSION 5.00
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmEnvioInfo 
   Caption         =   "Envio de info 10.00"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   8325
   Begin VB.Frame fraEnvioIncosistencias 
      Caption         =   "Envio Inconsistencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   180
      TabIndex        =   4
      Top             =   4740
      Width           =   7875
      Begin VB.CommandButton cmdIncosistencias 
         Caption         =   "Inconsistecias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5460
         TabIndex        =   12
         Top             =   1560
         Width           =   1200
      End
      Begin VB.TextBox txtPaso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   10
         Text            =   "C:\Envio\Inconsistencias"
         Top             =   1140
         Width           =   5175
      End
      Begin VB.TextBox txtDocumento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkNivelSolo 
         Caption         =   "Solo Este Nivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   7
         Top             =   720
         Width           =   1395
      End
      Begin Controles.cltGenerico ctlClienteIncon 
         Height          =   375
         Left            =   1500
         TabIndex        =   9
         Top             =   300
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
      End
      Begin VB.Label Label4 
         Caption         =   "Paso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Nº de Documento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraEnvioReferencia 
      Caption         =   "Envio Referencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7935
      Begin VB.CommandButton Command2 
         Caption         =   "robert"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   4200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkSoloDiccionario 
         Caption         =   "Enviar Solo DIcionario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   20
         Top             =   1920
         Width           =   3195
      End
      Begin VB.CheckBox chkEnvioTodoDocumentos 
         Caption         =   "Enviar todos los numero de documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3420
         TabIndex        =   19
         Top             =   1560
         Width           =   3195
      End
      Begin VB.CheckBox chkZip 
         Caption         =   "Archivo Zip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPredeterminado 
         Height          =   1215
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmEnvioInfo.frx":0000
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton cmdEnvioPorUsuario 
         Caption         =   "Preparar Envio por Cliente por usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   14
         Top             =   3780
         Width           =   3360
      End
      Begin Controles.ctlClienteUsuario ctlClienteUsuario 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   3360
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   556
      End
      Begin VB.CommandButton cmdPrepararENvio 
         Caption         =   "Preparar Envio por Cliente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1140
         TabIndex        =   3
         Top             =   2880
         Width           =   2400
      End
      Begin Controles.cltGenerico cltCliente 
         Height          =   375
         Left            =   1140
         TabIndex        =   1
         Top             =   2460
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
      End
      Begin VB.Label Label6 
         Caption         =   "Cuerpo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEnvioInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function SendMail(sEmailRecipient As String, sEmailSubject As String, sEmailBody As String, Optional sAttachment1 As String, Optional sAttachment2 As String, Optional sAttachment3 As String, Optional CopiaOculta As String)

'-----Send an Email Message using Outlook 98-----

'Developers Note:  In References, the Microsoft Outlook 98 Object Model must be selected for this to work
'                  Actual file is MSOUTL85.OLB

Dim emailOutlookApp As Outlook.Application
Dim emailNameSpace As Outlook.NameSpace
Dim emailFolder As Outlook.MAPIFolder
Dim emailItem As Outlook.MailItem
Dim EmailRecipient As Recipient

'-----Open Outlook in a background process and the Inbox Folder-----
Set emailOutlookApp = CreateObject("Outlook.Application")
Set emailNameSpace = emailOutlookApp.GetNamespace("MAPI")
Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderInbox)

'Enable the next line to actually see Outlook Open
'emailFolder.Display

'-----Create a new mail message, set the recipient, subject, and body-----
Set emailItem = emailOutlookApp.CreateItem(olMailItem)
Set EmailRecipient = emailItem.Recipients.Add(sEmailRecipient)
emailItem.Subject = sEmailSubject
If CopiaOculta <> "" Then
emailItem.BCC = CopiaOculta
End If
emailItem.Body = sEmailBody
If sAttachment1 <> "" Then
    emailItem.Attachments.Add sAttachment1
End If
If sAttachment2 <> "" Then
    emailItem.Attachments.Add sAttachment2
End If
If sAttachment3 <> "" Then
    emailItem.Attachments.Add sAttachment3
End If

'-----Send the Email-----
emailItem.Save
emailItem.Send

'-----Close the Outlook Application------
Rem emailOutlookApp.Quit

'----Inform User of Success-----
Rem MsgBox "Email was sent.", vbInformation

'-----Clear out the memory space held by variables-----
'Usually unnecessary but a good practice

Set emailNameSpace = Nothing
Set emailFolder = Nothing
Set emailItem = Nothing
Set emailOutlookApp = Nothing

End Function

Private Sub Check1_Click()

End Sub

Private Sub chkEnvio_Click()

End Sub

Private Sub cltCliente_Click()
    ctlClienteUsuario.LlenarConCliente cltCliente.Valor
End Sub


Private Sub cmdEnvioPorUsuario_Click()

Rem  Envio_Referencias cltCliente.Valor, ctlClienteUsuario.Valor



    If IsNull(cltCliente.Valor) Then
        MsgBox "Ingrese el cliente"
        Exit Sub
    End If
    If IsNull(ctlClienteUsuario.Valor) Then
        MsgBox "Ingrese el Usuario"
        Exit Sub
    End If

    Dim RsUsuario As New ADODB.Recordset
    Dim Sql As String
    Sql = "SELECT ID_CLIENTEUSUARIO, APELLIDO_NOMBRE, CORREO, Cod_Indice"
    Sql = Sql & vbCrLf & " From CLIENTEUSUARIO "
    Sql = Sql & vbCrLf & "  WHERE ID_CLIENTEUSUARIO = " & ctlClienteUsuario.Valor
    RsUsuario.Open Sql, ConActiva, 0, 1
     EnvioReferencias cltCliente.Valor, ctlClienteUsuario.Valor, Trim(RsUsuario!APELLIDO_NOMBRE), Trim(RsUsuario!correo), RsUsuario!Cod_Indice, True
    End Sub


Private Sub cmdIncosistencias_Click()
EnvioInconsistencias
End Sub

Private Sub EnvioReferencias(COD_CLIENTE As Integer, cod_Usuario As Integer, Apellido As String, correo As String, Cod_Indice As String, MostrarFinalizado As Boolean)

    Dim Sql As String
    Dim SqlIndice As String
    Dim rsbasa As New ADODB.Recordset
    Dim PasoPlanilla As String
    
   
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim ExcelIndice As Excel.Worksheet
    Dim ExcelReferenciaIndice As Excel.Worksheet
    Dim ExcelReferenciaCaja As Excel.Worksheet
    Rem COD_CLIENTE = cltCliente.Valor
    
    Dim SqlBase As String
    SqlBase = " SELECT  COD_ID_REFERENCIA, INDICES.TITULOHERENCIA,INDICES.DESCRIPCION AS DESCRIPCIONINDICE , REFERENCIAS.NRO_CAJA, REFERENCIAS.ITEM,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.INDICE, REFERENCIAS.DESCRIPCION,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.COD_CLIENTE,INDICES.ID_CODIGO_DOCUMENTO,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.EXPEDIENTE, REFERENCIAS.APELLIDO_NOMBRE, REFERENCIAS.BORRADO"
    SqlBase = SqlBase & vbCrLf & " From REFERENCIAS, INDICES"


   On Error GoTo er
       
       'Plantilla Base
       MousePointer = 11
       PasoPlanilla = "C:\Envio\" & Format(cod_Usuario, "0000") & "  " & Apellido & " " & Format(date, "dd_mm_yyyy") & "  Referencia.xls"
       If Dir(PasoPlanilla) <> "" Then
           Kill PasoPlanilla
       End If
       FileCopy strPasoPlanillas & "Referencia Envio.xls", PasoPlanilla
    
    'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(PasoPlanilla)
        Set ExcelIndice = libroEx.Worksheets.Item(1)
        Set ExcelReferenciaIndice = libroEx.Worksheets.Item(2)
        Set ExcelReferenciaCaja = libroEx.Worksheets.Item(3)
    
    Dim i As Integer
    For i = 1 To 10
    Rem MsgBox ExcelIndice.Cells(2, I).NumberFormat
    Next
    
    'Creacion del Indice
            Sql = " SELECT COD_CLIENTE, INDICE, ID_CODIGO_DOCUMENTO,"
            Sql = Sql & vbCrLf & "  Descripcion , Fecha, NUMERO, LETRA,TIPO_INDICE , HABILITAR_FECHA_DESDE , TITULO_FECHA_DESDE  ,HABILITAR_NRO_DESDE , TITULO_NRO_DESDE"
Sql = Sql & vbCrLf & "  ,HABILITAR_LETRA_DESDE, TITULO_LETRA_DESDE "
            Sql = Sql & vbCrLf & "  From INDICES "
            Sql = Sql & vbCrLf & "  WHERE COD_CLIENTE = " & COD_CLIENTE
            Sql = Sql & vbCrLf & " AND INDICE LIKE '" & Cod_Indice & "%'"
            Sql = Sql & vbCrLf & "  ORDER BY INDICE"
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, 0, 1
            ExcelIndice.Cells(1, 2) = " Planilla de referencia perteneciente a  " & Apellido
            IndiceExcel rsbasa, ExcelIndice
    
   If chkSoloDiccionario.value = 0 Then
    ' Creacion de referencia por Indice
    If Cod_Indice = "0" Then
             Sql = SqlBase & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
            Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
            Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & COD_CLIENTE
           
            Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.INDICE, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE"
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, 0, 1
            ExcelReferenciaIndice.Name = "Ref Indice"
            ProcesarPorIndices rsbasa, ExcelReferenciaIndice
            libroEx.Save
    Else
            Sql = SqlBase & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
            Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
            Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & COD_CLIENTE
            Sql = Sql & vbCrLf & " AND REFERENCIAS.INDICE LIKE '" & Cod_Indice & "%'"
            Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.INDICE, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE"
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, 0, 1
            
       
          Dim B As New Excel.QueryTable
          
          
           
          
Rem          ExcelReferenciaIndice.QueryTables.Add ConActiva, ExcelReferenciaIndice.Range("A1", "J2500"), Sql
          
            ExcelReferenciaIndice.Name = "Ref Indice"
            ProcesarPorIndices rsbasa, ExcelReferenciaIndice
            libroEx.Save
            End If
            
            
     End If
     
    If chkSoloDiccionario.value = 0 Then
   'Referencia por Caja
            Sql = SqlBase & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
            Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
            Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & COD_CLIENTE
            Sql = Sql & vbCrLf & " AND REFERENCIAS.INDICE LIKE '" & Cod_Indice & "%'"
            Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.NRO_CAJA, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE "
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, adOpenDynamic
            ProcesarPorCajas rsbasa, ExcelReferenciaCaja
            libroEx.Save
     End If

    rsbasa.Close
    libroEx.Save
    libroEx.Close
    ApExcel.Quit
    Set hojaEx = Nothing
    Set libroEx = Nothing
    Set ApExcel = Nothing
    
Dim ArchivoZip1 As String
    
'    If chkZip.value = 1 Then
'        ArchivoZip1 = Mid(PasoPlanilla, 1, CInt(Len(PasoPlanilla) - 4))
'        ZipArchivo PasoPlanilla, ArchivoZip1
'        SendMail correo, "Envio de Referencias " & Format(date, "dd/mm/2006"), txtPredeterminado.Text, ArchivoZip1 & ".zip"
'     Else
'        If FileLen(PasoPlanilla) > 1500000 Then
'            ArchivoZip1 = Mid(PasoPlanilla, 1, CInt(Len(PasoPlanilla) - 4))
'            ZipArchivo PasoPlanilla, ArchivoZip1
'            Rem SendMail correo, "Envio de Referencias " & Format(date, "dd/mm/2006"), txtPredeterminado.Text, ArchivoZip1 & ".zip"
'        Else
'           Rem  SendMail correo, "Envio de Referencias " & Format(date, "dd/mm/2006"), txtPredeterminado.Text, PasoPlanilla
'        End If
'    End If
    MousePointer = 0
    If MostrarFinalizado = True Then
        MsgBox "Se relizo con Exito", vbInformation
    End If
    Exit Sub
er:
    If Err.Number = 287 Then
      MousePointer = 0
        Exit Sub
        MousePointer = 0
    End If
    MousePointer = 0
    MsgBox Err.Description
   
     rsbasa.Close
    libroEx.Save
 
    libroEx.Close
    ApExcel.Quit
    Set hojaEx = Nothing
    Set libroEx = Nothing
    Set ApExcel = Nothing
End Sub


Private Sub ctlClienteUsuario1_SectorEncontrado(Sector As String)

End Sub

Private Sub cmdPrepararENvio_Click()

If MsgBox("Envio por excel ", vbYesNo) = vbYes Then

    Dim RsUsuario As New ADODB.Recordset
    Dim Sql As String




Sql = "  SELECT     REMITOS_CUERPO.COD_USUARIO_CLIENTE, CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.APELLIDO_NOMBRE, CLIENTEUSUARIO.COD_INDICE,"
Sql = Sql & vbCrLf & "                      CLIENTEUSUARIO.correo"
Sql = Sql & vbCrLf & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & vbCrLf & "                       CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
Sql = Sql & vbCrLf & " WHERE     (REMITOS_CUERPO.TIPO = 0) AND (REMITOS_CUERPO.ID_CLIENTE = 4) AND (REMITOS_CUERPO.FECHA > '" & InputBox("Ingrese la fecha desde", "Fecha", Format(Now, "DD/MM/YYYY")) & "'"
Sql = Sql & vbCrLf & ")  GROUP BY REMITOS_CUERPO.COD_USUARIO_CLIENTE, CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.APELLIDO_NOMBRE, CLIENTEUSUARIO.COD_INDICE,"
Sql = Sql & vbCrLf & "                       CLIENTEUSUARIO.correo"

    Sql = "  SELECT     REMITOS_CUERPO.COD_USUARIO_CLIENTE, CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.APELLIDO_NOMBRE, CLIENTEUSUARIO.COD_INDICE,"
Sql = Sql & vbCrLf & "                      CLIENTEUSUARIO.correo"
Sql = Sql & vbCrLf & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & vbCrLf & "                       CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
Sql = Sql & vbCrLf & " WHERE      (REMITOS_CUERPO.ID_CLIENTE = 4)"
Sql = Sql & vbCrLf & "  GROUP BY REMITOS_CUERPO.COD_USUARIO_CLIENTE, CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.APELLIDO_NOMBRE, CLIENTEUSUARIO.COD_INDICE,"
Sql = Sql & vbCrLf & "                       CLIENTEUSUARIO.correo"
    
    Dim directorio As String
    Dim PasoFinal As String
    Diretorio = Format(Now, "ddmmyyyy HHMM")
    If Dir("C:\EnvioInfo\*.*", vbDirectory) = "" Then
        FileSystem.MkDir "C:\EnvioInfo"
        FileSystem.MkDir "C:\EnvioInfo\" & Diretorio
    Else
        If Dir("C:\EnvioInfo\" & Diretorio, vbDirectory) = "" Then
            FileSystem.MkDir "C:\EnvioInfo\" & Diretorio
        End If
        
    End If
    RsUsuario.Open Sql, strConBasa
    
    Do While Not RsUsuario.EOF
    
    
        EnvioReferencias 4, RsUsuario!ID_CLIENTEUSUARIO, RsUsuario!APELLIDO_NOMBRE, "", RsUsuario!Cod_Indice, False
        RsUsuario.MoveNext
    Loop
    
    Else
    
Envio_Referencias cltCliente.Valor, 0
End If
End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset

Dim Sql As String
Dim correo As String

correo = "Estimados Clientes:"
correo = correo & vbCrLf & " Nos dirigimos a Ustedes para informarles que el día 10/08/2012 realizaremos las tareas habituales en el horario de 8 a 14 horas debido a cambios en nuestro sistema informático por lo que agradeceríamos que los pedidos fueran realizados con anticipación. Las entrega de consultas de la tarde serán reprogramadas para el día lunes en la mañana"
correo = correo & vbCrLf & "Agradecemos su colaboración."
correo = correo & vbCrLf & "     Saludos cordiales."



Sql = " SELECT     APELLIDO_NOMBRE, CORREO, DESHABILITADO, ID_CLIENTEUSUARIO"
Sql = Sql & " From CLIENTEUSUARIO"
Sql = Sql & " Where (DESHABILITADO Is Null) And (Not (correo Is Null))"
Sql = Sql & " ORDER BY CORREO"
rs.Open Sql, ConActiva


Dim i As Integer
Dim CC As String

Do While Not rs.EOF

 i = i + 1
If i > 1500 Then
     SendMail rs!correo, "NOTIFICACION ARCHIVO", correo

    Rem CC = CC & ";" &    End If
   End If
    rs.MoveNext
Loop




End Sub

Private Sub Command2_Click()
    EnvioReferencias 4, 186, "ROBERT HUGO", "", "0", True
End Sub

Private Sub ctlClienteUsuario_SectorEncontrado(Sector As String)
frmEnvioInfo.Refresh
End Sub

Private Sub Form_Load()
    cltCliente.TipoControl = Cliente
    ctlClienteIncon.TipoControl = Cliente
End Sub




Public Sub EnvioInconsistencias()
    Dim Sql As String
    Dim SqlIndice As String
    Dim rsbasa As New ADODB.Recordset
    Dim RefIndice As String
    Dim PlanillaInc As String
   
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim ExcelIndice As Excel.Worksheet
    Dim ExcelReferenciaIndice As Excel.Worksheet
    Dim ExcelReferenciaCaja As Excel.Worksheet
    
    
    Dim SqlBase As String
    SqlBase = " SELECT  COD_ID_REFERENCIA, INDICES.TITULOHERENCIA,INDICES.DESCRIPCION AS DESCRIPCIONINDICE , REFERENCIAS.NRO_CAJA, REFERENCIAS.ITEM,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.INDICE, REFERENCIAS.DESCRIPCION,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.COD_CLIENTE,INDICES.ID_CODIGO_DOCUMENTO,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.EXPEDIENTE, REFERENCIAS.APELLIDO_NOMBRE, REFERENCIAS.BORRADO"
    SqlBase = SqlBase & vbCrLf & " From REFERENCIAS, INDICES"


   On Error GoTo er
   
   
   If IsNull(ctlClienteIncon.Valor) Then
    Exit Sub
   End If
   RefIndice = TraerIncide(ctlClienteIncon.Valor, txtDocumento.Text)
   If RefIndice = "" Then
        MsgBox "Error en documento"
        Exit Sub
    End If
   
   
       
       'Plantilla Base
       MousePointer = 11
       PlanillaInc = txtPaso.Text & "\DOC " & Format(txtDocumento.Text, "0000") & "   " & Format(date, "dd_mm_yyyy") & "  Corregir.xls"
       If Dir(PlanillaInc) <> "" Then
           Kill PlanillaInc
       End If
       FileCopy "strPasoPlanillasReferencia Envio.xls", PlanillaInc
    
    'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(PlanillaInc)
        Set ExcelIndice = libroEx.Worksheets.Item(1)
        Set ExcelReferenciaIndice = libroEx.Worksheets.Item(2)
        Set ExcelReferenciaCaja = libroEx.Worksheets.Item(3)
        
    
    
    'Creacion del Indice
            Sql = " SELECT COD_CLIENTE, INDICE, ID_CODIGO_DOCUMENTO,"
            Sql = Sql & vbCrLf & "  Descripcion , Fecha, NUMERO, LETRA"
            Sql = Sql & vbCrLf & "  From INDICES "
            Sql = Sql & vbCrLf & "  WHERE COD_CLIENTE = " & ctlClienteIncon.Valor
            Sql = Sql & vbCrLf & " AND INDICE LIKE '" & RefIndice & "%'"
            Sql = Sql & vbCrLf & "  ORDER BY INDICE"
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, 0, 1
            ExcelIndice.Cells(2, 4) = " Diccionario de Documentos "
            IndiceExcel rsbasa, ExcelIndice
    ' Creacion de referencia por Indice
            
 
        Sql = SqlBase & vbCrLf & "  WHERE REFERENCIAS.INDICE = INDICES.INDICE AND"
        Sql = Sql & vbCrLf & " REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
        Sql = Sql & vbCrLf & " REFERENCIAS.COD_CLIENTE = " & ctlClienteIncon.Valor
        If chkNivelSolo.value = 1 Then
            Sql = Sql & vbCrLf & " AND REFERENCIAS.INDICE LIKE '" & RefIndice & "'"
        Else
            Sql = Sql & vbCrLf & " AND REFERENCIAS.INDICE LIKE '" & RefIndice & "%'"
        End If
        Sql = Sql & vbCrLf & " AND INDICES.TIPO_INDICE <> 'Documento'"
        Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.INDICE, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE"
        Set rsbasa = New ADODB.Recordset
        rsbasa.Open Sql, ConActiva, 0, 1
        ExcelReferenciaIndice.Name = "Inconsistencias"
        ExcelReferenciaIndice.Cells(1, 3) = "Planilla de Inconsistencias"
        Rem ProcesarPorIndices rsbasa, ExcelReferenciaIndice
   'Referencia por Caja
        Rem ExcelReferenciaCaja.Visible = False
 

    rsbasa.Close
    libroEx.Save
    libroEx.Close
    ApExcel.Quit
    Set hojaEx = Nothing
    Set libroEx = Nothing
    Set ApExcel = Nothing
    MousePointer = 0
        MsgBox "La exportacion a finalizado", vbInformation
    Exit Sub
er:
    MousePointer = 0
    MsgBox Err.Description
    libroEx.Close
    ApExcel.Quit
    Set hojaEx = Nothing
    Set libroEx = Nothing
    Set ApExcel = Nothing
    MsgBox "Exportacion terminada"
End Sub


Public Function TraerIncide(COD_CLIENTE As Integer, Doc As Integer) As String
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Sql = " SELECT INDICE "
    Sql = Sql & vbCrLf & " From INDICES "
    Sql = Sql & vbCrLf & " Where ID_CODIGO_DOCUMENTO = " & Doc
    Sql = Sql & vbCrLf & " And cod_Cliente = " & COD_CLIENTE
    rs.Open Sql, ConActiva, 0, 1
    If Not rs.EOF Then
        TraerIncide = rs!Indice
    Else
        TraerIncide = ""
    End If
    

End Function



Public Sub Envio_Referencias(Cliente As Integer, Usuario As Integer)
If IsNull(cltCliente.Valor) Then
        MsgBox "Ingrese el cliente"
        Exit Sub
    End If
    
    MousePointer = 11
    TituloHerencia Cliente
    Dim RsUsuario As New ADODB.Recordset
    Dim Sql As String
    Sql = "SELECT ID_CLIENTEUSUARIO, APELLIDO_NOMBRE, CORREO, Cod_Indice"
    Sql = Sql & vbCrLf & " From CLIENTEUSUARIO "
    Sql = Sql & vbCrLf & " WHERE cod_cliente = " & Cliente
    If Usuario = 0 Then
    Sql = Sql & vbCrLf & " and REFERENCIAS = '1'"
    Rem Sql = Sql & vbCrLf & " AND NOT (CORREO IS NULL)"
    Sql = Sql & vbCrLf & " AND NOT (COD_INDICE IS NULL) "
    Else
     Sql = Sql & vbCrLf & " and  ID_CLIENTEUSUARIO = " & Usuario
    End If
   Rem  Sql = Sql & vbCrLf & " AND (FECHA_ENVIO_REFERENCIAS IS NULL)"
   
   Rem     sql = sql & vbCrLf & " and  (FECHA_ENVIO_REFERENCIAS <> CONVERT(DATETIME, '2010-02-08 00:00:00', 102))"
   
    Sql = Sql & vbCrLf & " ORDER BY ID_CLIENTEUSUARIO"
    Dim directorio As String
    Dim PasoFinal As String
    Diretorio = Format(Now, "ddmmyyyy HHMM")
    If Dir("C:\EnvioInfo\*.*", vbDirectory) = "" Then
        FileSystem.MkDir "C:\EnvioInfo"
        FileSystem.MkDir "C:\EnvioInfo\" & Diretorio
    Else
        If Dir("C:\EnvioInfo\" & Diretorio, vbDirectory) = "" Then
            FileSystem.MkDir "C:\EnvioInfo\" & Diretorio
        End If
        
    End If
    RsUsuario.Open Sql, strConBasa
    Dim Cod_Indice As String
    Do While Not RsUsuario.EOF
    
        PasoFinal = "C:\EnvioInfo\" & Diretorio & "\" & RsUsuario!APELLIDO_NOMBRE & " Diccionario.xls"
        
        Cod_Indice = RsUsuario!Cod_Indice
        Cod_Indice = "001001"
        
        Sql = " SELECT * "
        Sql = Sql & " FROM   V_INDICES "
        Sql = Sql & " Where COD_CLIENTE = " & Cliente
        Sql = Sql & " AND (INDICE LIKE '" & Cod_Indice & "%')"
        Sql = Sql & "  ORDER BY INDICE "
        frmReportes.ExportarReporte PasoReportes & "rptIndicesNuevos.rpt", Sql, True, "", "", PasoFinal
Rem frmReportes.ImprimirReporte PasoReportes & "rptIndicesNuevos.rpt", Sql, True, "", ""
        
        PasoFinal = "C:\EnvioInfo\" & Diretorio & "\" & RsUsuario!APELLIDO_NOMBRE & " Referencias.xls"
        Sql = " SELECT     * From basasql.dbo.V_REFERENCIAS"
        Sql = Sql & " Where COD_CLIENTE = " & cltCliente.Valor
        Sql = Sql & " AND (INDICE LIKE '" & Cod_Indice & "%')"
        Sql = Sql & " ORDER BY INDICE "
        frmReportes.ExportarReporte PasoReportes & "rptreferencias.rpt", Sql, True, "", "", PasoFinal
        RsUsuario.MoveNext
   Loop
   MousePointer = 0
   MsgBox "Envio de referencia  completo", vbInformation


End Sub

