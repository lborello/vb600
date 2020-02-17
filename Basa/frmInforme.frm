VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInforme 
   Caption         =   "Informe"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "Exportar Excel"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   4380
      Width           =   2475
   End
   Begin MSDataGridLib.DataGrid grdInforme 
      Height          =   3555
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   6271
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
   Begin VB.Label LblTituloInforme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7335
   End
End
Attribute VB_Name = "frmInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CargarInforme(Titulo As String, rs As ADODB.Recordset)
    

    LblTituloInforme.Caption = Titulo
    DATOSGRILLA grdInforme, rs
End Sub

Private Sub cmdExportarExcel_Click()
 CopiarDatosGrilla grdInforme
End Sub

Private Sub Form_Resize()
On Error GoTo Salir
grdInforme.Height = frmInforme.Height - 2500

grdInforme.Width = frmInforme.Width - 150
cmdExportarExcel.Top = grdInforme.Height + 800
Salir:
End Sub
