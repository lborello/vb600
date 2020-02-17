VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmEnvioCorreo 
   Caption         =   "CORREO"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11505
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
   ScaleHeight     =   7305
   ScaleWidth      =   11505
   Begin VB.TextBox txtCorreo 
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
      Left            =   960
      TabIndex        =   9
      Top             =   4140
      Width           =   4935
   End
   Begin Controles.ctlClienteUsuario ctlClienteUsuario1 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   4920
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   661
   End
   Begin VB.TextBox txtAdjunto4 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Text            =   "Z:\Administracion\Referencias\Planilla Modelo.xls"
      Top             =   6840
      Width           =   6295
   End
   Begin VB.TextBox txtAdjunto3 
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Text            =   "Z:\Administracion\Referencias\Registro de usuario.pdf"
      Top             =   6360
      Width           =   6295
   End
   Begin VB.TextBox txtAdjunto2 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Text            =   "Z:\Administracion\Referencias\Referencias Manuales.pdf"
      Top             =   5880
      Width           =   6295
   End
   Begin VB.TextBox txtAdjunto1 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Text            =   "Z:\Administracion\Referencias\Procedimiento de retiro de cajas.pdf"
      Top             =   5400
      Width           =   6295
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4560
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdEnvioCorreo 
      Caption         =   "Enviar Correo"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtSuject 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "CAMBIO DE METODOLOGIA EN RETIRO Y REFERENCIACION DE CAJAS"
      Top             =   120
      Width           =   6975
   End
   Begin RichTextLib.RichTextBox txtCuerpo 
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmEnvioCorreos.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Correo:"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
End
Attribute VB_Name = "frmEnvioCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdEnvioCorreo_Click()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim correo As String
        Dim Indice As String
        Dim i As Integer
        
        If txtCorreo.Text = "" Then
                Sql = " SELECT     CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.COD_CLIENTE, CLIENTEUSUARIO.APELLIDO_NOMBRE, CLIENTEUSUARIO.DESHABILITADO,"
                Sql = Sql & vbCrLf & " CLIENTEUSUARIO.correo , CLIENTEUSUARIO.FECHA_ENVIO_REFERENCIAS, INDICES.ID_CODIGO_DOCUMENTO, INDICES.Descripcion"
                Sql = Sql & vbCrLf & " FROM CLIENTEUSUARIO INNER JOIN "
                Sql = Sql & vbCrLf & " INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE"
                Sql = Sql & vbCrLf & " Where (CLIENTEUSUARIO.DESHABILITADO Is Null)"
                Sql = Sql & vbCrLf & " And CLIENTEUSUARIO.COD_CLIENTE = " & ctlCliente.Valor
                If Not IsNull(ctlClienteUsuario1.Valor) Then
                     Sql = Sql & vbCrLf & " And CLIENTEUSUARIO.ID_CLIENTEUSUARIO = " & ctlClienteUsuario1.Valor
                     ctlClienteUsuario1.Valor = 0
                End If
                Rem sql = sql & vbCrLf & " And CLIENTEUSUARIO.FECHA_ENVIO_REFERENCIAS < " & FechaFormato(Now)
                Sql = Sql & vbCrLf & " ORDER BY CLIENTEUSUARIO.COD_CLIENTE"
                rs.Open Sql, ConActiva
                Do While Not rs.EOF
                    i = i + 1
                    Indice = "Estimado/da: " & UCase(rs!APELLIDO_NOMBRE)
                    Indice = Indice & vbCrLf & " Tenemos registrado en nuestra base de datos que usted  pertenece al sector : " & UCase(rs!Descripcion)
                    Indice = Indice & vbCrLf & " cuyo numero de Indice Interno es : " & rs!ID_CODIGO_DOCUMENTO
                    Indice = Indice & vbCrLf & " En el caso de no ser correcto el sector / Sucursal le solicitamos que "
                    Indice = Indice & vbCrLf & " nos envie por este medio su sucursal / sector Actual "
                    Indice = Indice & vbCrLf & " Gracias "
                    EnvioCorreo Trim(rs!correo), txtSuject.Text, Indice & vbCrLf & txtCuerpo.Text, txtAdjunto1.Text, txtAdjunto2.Text, txtAdjunto3.Text, txtAdjunto4.Text
                    lblCOrreoEnviados.Caption = i
                    lblCOrreoEnviados.Refresh
                    Sql = " UPDATE    CLIENTEUSUARIO "
                    Sql = Sql & " SET   FECHA_ENVIO_REFERENCIAS = " & SysDate
                    Sql = Sql & " Where ID_CLIENTEUSUARIO = " & rs!ID_CLIENTEUSUARIO
                    ExecutarSql Sql
                    rs.MoveNext
                Loop
        Else
                EnvioCorreo txtCorreo.Text, txtSuject.Text, Indice & vbCrLf & txtCuerpo.Text, txtAdjunto1.Text, txtAdjunto2.Text, txtAdjunto3.Text, txtAdjunto4.Text
        
        End If
        
        
        MsgBox "Terminado"

End Sub

Public Function EnvioCorreo(sEmailRecipient As String, sEmailSubject As String, sEmailBody As String, Optional sAttachment1 As String, Optional sAttachment2 As String, Optional sAttachment3 As String, Optional sAttachment4 As String, Optional CopiaOculta As String)

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
If sAttachment4 <> "" Then
    emailItem.Attachments.Add sAttachment4
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

Private Sub ctlCliente_Click()
ctlClienteUsuario1.LlenarConCliente (ctlCliente.Valor)
End Sub

Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
    Dim t As String
    
  t = "En búsqueda de una constante mejora del servicio se ha decidido un cambio metodológico con respecto al envío de cajas para obtener mejor control de la documentación enviada a nuestras instalaciones."
t = t & vbCrLf & "    Usted recibirá adjunto a éste correo 4 archivos:"
t = t & vbCrLf
t = t & vbCrLf & "    1-Procedimiento de retiro de caja.pdf - Este archivo es una breve descripción de cómo se deberán enviar las cajas y las referencias de las mismas."
t = t & vbCrLf
t = t & vbCrLf & "    2-Registro de usuario.pdf - Este archivo solicita los datos de usuario para completar o verificar en nuestros sistemas. Le solicitamos que por favor se complete el mismo y lo tengan disponible para cuando el personal de nuestra empresa lo visite."
t = t & vbCrLf
t = t & vbCrLf & "    3-Referencias Manuales.pdf - Es el formato de la planilla de referencias, el cual deberá se impreso en hoja A-4 o solicitar al personal de nuestra empresa las planillas correspondientes."
t = t & vbCrLf
t = t & vbCrLf & "    4-Planilla Modelo.xls - Este archivo (Excel) es la planilla modelo para el envío de referencias a través de correo electrónico que deberá ser enviada a la siguiente dirección de correo electrónico: analistadocumental@basaargentina.com.ar <mailto:analistadocumental@basaargentina.com.ar> <<mailto:analistadocumental@basaargentina.com.ar>>"
t = t & vbCrLf
t = t & vbCrLf & "Desde ya muchas gracias por colaborar con la mejora constante del servicio."
t = t & vbCrLf & " Ante cualquier  consulta: 4217616 / 4217618   Interno: 12  - 16 - 11 "
t = t & vbCrLf & " Banco de Archivos s.a"
txtCuerpo.Text = t
    
End Sub


