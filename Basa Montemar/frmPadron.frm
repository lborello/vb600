VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPadron 
   Caption         =   "Padron"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
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
   ScaleHeight     =   6900
   ScaleWidth      =   10065
   Begin MSDataGridLib.DataGrid grdPadron 
      Height          =   5955
      Left            =   180
      TabIndex        =   5
      Top             =   780
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10504
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   180
      Width           =   1395
   End
   Begin VB.TextBox txtApellidoNombre 
      Height          =   375
      Left            =   5220
      TabIndex        =   3
      Top             =   180
      Width           =   2775
   End
   Begin VB.TextBox txtDocumento 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Apellido Nombre"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Documento:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmPadron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
Dim con As New ADODB.Connection
    Dim BackColor As ColorConstants
    con.Open strConBasa
    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    If txtDocumento.Text <> "" Then
      rs.Open "SELECT DOCUMENTO , APELLIDO_NOMBRE From PADRON Where DOCUMENTO = " & txtDocumento.Text, con
    Else
     If txtApellidoNombre.Text <> "" Then
        rs.Open "SELECT DOCUMENTO , APELLIDO_NOMBRE From PADRON Where APELLIDO_NOMBRE  like  '" & txtApellidoNombre.Text & "%'", con
     
     End If
     
    End If
    
   Set grdPadron.DataSource = rs.DataSource
   grdPadron.Refresh
    
    

End Sub
