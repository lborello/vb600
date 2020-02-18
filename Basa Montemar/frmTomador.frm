VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTomador 
   Caption         =   "Tomador"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   8520
   Begin MSDataGridLib.DataGrid grdTomador 
      Height          =   3555
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6271
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
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   720
      Width           =   1155
   End
   Begin VB.TextBox txtTomador 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   300
      Width           =   6795
   End
   Begin VB.Label Label1 
      Caption         =   "Tomador"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmTomador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTomador As ADODB.Recordset

Private Sub cmdNuevo_Click()


If Trim(UCase(txtTomador.Text)) <> "" Then
Dim rs As New ADODB.Recordset
rs.Open "SELECT     MAX( ID_TOMADOR) as MaxID FROM         basasql.dbo.LA_CAJA_TOMADOR", ConActiva

    ExecutarSql "  INSERT INTO LA_CAJA_TOMADOR  (ID_TOMADOR  , DESCRIPCION) VALUES ( " & rs!MaxID + 1 & ",  '" & Trim(UCase(txtTomador.Text)) & "' )"
    RefrescarTomador
    rsTomador.Filter = "DESCRIPCION LIKE '%" & Trim(UCase(txtTomador.Text)) & "%'"
    txtTomador.Text = ""
    End If
End Sub

Private Sub Form_Load()
RefrescarTomador
grdTomador.Columns.Item(0).Width = 1500
grdTomador.Columns.Item(1).Width = 5000
End Sub
Private Sub txtTomador_Change()
If Len(txtTomador.Text) > 3 Then
    rsTomador.Filter = "DESCRIPCION LIKE '%" & txtTomador.Text & "%'"
Else
    rsTomador.Filter = ""
    rsTomador.Requery
End If

End Sub

Public Sub RefrescarTomador()
 Set rsTomador = New ADODB.Recordset
 rsTomador.CursorLocation = adUseClient
 Dim Sql As String
 Sql = " SELECT     ID_TOMADOR, DESCRIPCION"
Sql = Sql & " From LA_CAJA_TOMADOR"
Sql = Sql & "  ORDER BY DESCRIPCION"
    
    rsTomador.Open Sql, ConActiva, 0, 1
 Set grdTomador.DataSource = rsTomador
 
End Sub
