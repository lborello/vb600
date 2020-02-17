VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContronReferencias 
   Caption         =   "Control de referencia"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form4"
   ScaleHeight     =   8730
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtCajaControlReferencia 
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3075
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7635
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   13467
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
            LCID            =   11274
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
            LCID            =   11274
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
End
Attribute VB_Name = "frmContronReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim strConAsp As String
    Dim strConAsp150 As String
    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    sql = " SELECT     Caja.Id, Caja.Numero, Caja.CajaAzul, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1, Documento.TEXTO2, Documento.FECHA1,"
    sql = sql & " Documento.fecha2 , caja.CAJA_ASP"
    sql = sql & "  FROM         Caja INNER JOIN"
    sql = sql & " Documento ON Caja.Id = Documento.IdCaja"
    sql = sql & "  WHERE     (Caja.Numero LIKE N'%" & txtCajaControlReferencia.Text & "%')"
    rs.CursorLocation = adUseClient
    rs.Open sql, strConAsp150
    Set DataGrid1.DataSource = rs.DataSource
    DataGrid1.Refresh



End Sub
