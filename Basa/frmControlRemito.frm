VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControlRemito 
   Caption         =   "Control Remito"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   9555
   Begin MSDataGridLib.DataGrid grdRemitoDetalle 
      Height          =   2955
      Left            =   180
      TabIndex        =   4
      Top             =   4740
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   5212
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   180
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txtFiltro 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   5535
   End
   Begin MSDataGridLib.DataGrid grdRemito 
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6482
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
End
Attribute VB_Name = "frmControlRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset

Private Sub Command1_Click()



Dim sql As String
Set rs = New ADODB.Recordset

sql = " SELECT     NRO_REMITO, NRO_REM_PROV, TIPO, OPERACION, ESTADO, FECHA, ID_CLIENTE, OBSERVACIONES, CANTIDAD, ANULADO, COBRAR_FLETE"
sql = sql & " From REMITOS_CUERPO "
sql = sql & " WHERE NRO_REMITO IN (" & txtFiltro.Text & ")"
rs.CursorLocation = adUseClient
rs.Open sql, strConBasa, 2, 3


Set grdRemito.DataSource = rs.DataSource
grdRemito.Refresh

End Sub

Private Sub Command2_Click()
Dim A As String
A = Trim(Clipboard.GetText)

A = Replace(A, vbCrLf, ",")
A = Replace(A, " ", "")
A = Mid(A, 1, Len(A) - 1)
txtFiltro.Text = A
End Sub

Private Sub Form_Resize()
 On Error GoTo salir
   grdRemito.Width = frmControlRemito.Width - 200
 
 
salir:
 
End Sub

Private Sub grdRemito_Click()
Dim rs As New ADODB.Recordset


Dim sql As String


sql = " SELECT     NRO_REMITO, NRO_CAJA, DESDE, HASTA, TIPO_ALMACENADO"
sql = sql & " From REMITOS_DETALLE"
sql = sql & " where NRO_REMITO in (" & grdRemito.Columns("NRO_REMITO").Text & ")"
sql = sql & " ORDER BY DESDE"


rs.CursorLocation = adUseClient
rs.Open sql, strConBasa, 2, 3


Set grdRemitoDetalle.DataSource = rs.DataSource

grdRemitoDetalle.Refresh

End Sub

Private Sub grdRemito_DblClick()

End Sub

Private Sub grdRemito_HeadClick(ByVal ColIndex As Integer)

rs.Sort = grdRemito.Columns(ColIndex).DataField
End Sub
