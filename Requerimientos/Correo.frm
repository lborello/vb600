VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCorreo 
   Caption         =   "Correo"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   9120
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   300
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1995
      Left            =   60
      TabIndex        =   1
      Top             =   3420
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3519
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1635
      Left            =   60
      TabIndex        =   0
      Top             =   1680
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   2884
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
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
      Caption         =   "Label2"
      Height          =   615
      Left            =   4620
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1260
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim emailOutlookApp As Outlook.Application
Dim emailNameSpace As Outlook.Namespace
Dim emailFolder As Outlook.MAPIFolder
Dim SubFolder  As Outlook.MAPIFolder
Dim emailItem As Outlook.MailItem
Dim EmailRecipient As Recipient
Dim emailItem2 As Outlook.MailItem

Rem -----Open Outlook in a background process and the Inbox Folder-----
Set emailOutlookApp = CreateObject("Outlook.Application")
Set emailNameSpace = emailOutlookApp.GetNamespace("MAPI")
 Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderInbox)
 
Rem Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderSentMail)

MsgBox emailFolder.Folders.Count
Dim i As Integer
Dim c As Integer


MsgBox emailFolder.Items.Count
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ENTRYID As String


For i = 1 To emailFolder.Items.Count
    Set emailItem2 = emailFolder.Items(i)

     emailItem2.FlagRequest = i

Next
    ProgressBar1.Max = emailFolder.Items.Count
    Dim F As Integer
   For F = 1 To emailFolder.Folders.Count - 1
   
          Label1.Caption = emailFolder.Folders.Item(F)
           Set SubFolder = emailFolder.Folders.Item(F)
           Label2.Caption = SubFolder.Items.Count
            For c = 1 To SubFolder.Items.Count - 1
                On Error GoTo saltar
                Rem if emailFolder.Items(c)
                If SubFolder.Items(c).MessageClass = "IPM.Note" And Mid(emailFolder.Items(c), 1, 14) <> "Sin entregar: " And Mid(emailFolder.Items(c), 1, 19) <> "Solicitud de tarea:" Then
                  
                    Set emailItem2 = emailFolder.Items(c)
                    
                ProgressBar1.Value = c
                
                        If emailItem2.SenderName <> "" Then
                Sql = " INSERT INTO CORREOS"
                Sql = Sql & "(NOMBRE, DIRECCION, DOMINIO, PERSONAL)"
                Sql = Sql & " VALUES     ('" & Replace(Trim(emailItem2.SenderName), "'", "") & "','" & Replace(Trim(emailItem2.SenderEmailAddress), "$", "") & "','" & Mid(emailItem2.SenderEmailAddress, InStr(1, Replace(Trim(emailItem2.SenderEmailAddress), "$", "@"), "@")) & "','" & MDIfrmInicio.StaInicio.Panels.Item(3).Text & "')"
                    End If
                         ExecutarSql Sql
saltar:
        
                 
                End If
        Next

Next

MsgBox "Listo"





'For c = 1 To emailFolder.Folders.Count
'
'        Set emailItem2 = emailFolder.Items(c)
'        MsgBox emailItem2.SenderName
'        MsgBox emailItem2.SenderEmailAddress
'
'
'   MsgBox emailFolder.Folders.Item(c).Name
'
'    Set emailItem2 = emailFolder.Items(c)
'    For i = 1 To emailFolder.Folders.Item(c).Items.Count
'    Set emailFolder = emailFolder.Folders.Item(c)
'        Set emailItem2 = emailFolder.Items(c)
'        MsgBox emailItem2.SenderName
'
'       MsgBox emailItem2.SenderEmailAddress
'     Next
'Next
'
''
''     Rem If Not IsNumeric(emailItem2.Categories)  Then
''     If Trim(emailItem2.Categories) = "" Then
''
''        Rem ENTRYID = Replace(emailItem2.InternetCodepage & emailItem2.ReceivedByEntryID & emailItem2.Subject, "'", "´" & Mid(emailItem2.Body, 1, 7000))
''        ENTRYID = emailItem2.ENTRYID
''
''           emailItem2.Categories = MaxCorreo
''
''        Sql = "  SELECT     ENTRYID, ID_CORREO"
''        Sql = Sql & " From dbo.CORREOS"
''        Sql = Sql & "  WHERE     (ENTRYID = '" & ENTRYID & " ')"
''        Set Rs = New ADODB.Recordset
''        Rs.Open Sql, strConBasa
''        If Rs.EOF Then
''           Sql = " INSERT INTO dbo.CORREOS"
''           Sql = Sql & " (ENTRYID, ENVIADO, ASUNTO, CUERPO, KF_USUARIO)"
''           Sql = Sql & " VALUES ('" & ENTRYID & "','" & emailItem2.SenderEmailAddress & "','" & Replace(emailItem2.Subject, "'", "´") & "','" & Mid(Replace(emailItem2.Body, "'", "`"), 1, 7000) & "'," & 99 & ")"
''           ExecutarSql Sql
''           emailItem2.Categories = MaxCorreo
''           emailItem2.Save
''        Else
''            emailItem2.Categories = Rs!ID_CORREO
''            emailItem2.Save
''
''        End If
''     End If
''
''Next
''
''For i = 1 To emailFolder.Folders.Item("SUPER").Items.Count
''    Set emailItem2 = emailFolder.Folders.Item("SUPER").Items(i)
''   Rem  MsgBox emailItem2.Body
''    emailItem2.Categories = "REQUE:" & i & " id " & emailItem2.ENTRYID
''    Rem emailItem2.FlagStatus = olFlagComplete
''
''     emailItem2.FlagRequest = i
''
''     Sql = " INSERT INTO dbo.CORREOS"
''      Sql = Sql & " (ENTRYID, ENVIADO, ASUNTO, CUERPO, KF_USUARIO)"
''Sql = Sql & " VALUES ('" & emailItem2.InternetCodepage & emailItem2.SenderEmailAddress & emailItem2.Subject & "','" & emailItem2.SenderEmailAddress & "','" & emailItem2.Subject & "','" & Replace(Mid(emailItem2.Body, 1, 2000), "'", "`") & "'," & ctlPersonal.Valor & ")"
''         ExecutarSql Sql
''
''    emailItem2.Save
''Next
'

Set emailNameSpace = Nothing
Set emailFolder = Nothing
Set emailItem = Nothing
Set emailOutlookApp = Nothing

MsgBox "ok"
End Sub

Public Function MaxCorreo() As Long
    Dim Rs As New ADODB.Recordset
    Rs.Open "SELECT     MAX(ID_CORREO) AS MaxCorreo FROM  dbo.CORREOS ", strConBasa
    
    MaxCorreo = Rs!MaxCorreo
    
    


End Function

