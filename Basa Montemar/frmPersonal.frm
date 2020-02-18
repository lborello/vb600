VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPersonal 
   Caption         =   "Personal"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13140
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
   ScaleHeight     =   8430
   ScaleWidth      =   13140
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   435
      Left            =   7380
      TabIndex        =   3
      Top             =   180
      Width           =   1335
   End
   Begin VB.CommandButton cmdcopiarexcel 
      Caption         =   "Copiar Excel"
      Height          =   435
      Left            =   8940
      TabIndex        =   2
      Top             =   180
      Width           =   1455
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   435
      Left            =   10620
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdPersonal 
      Height          =   7515
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   13256
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      AllowAddNew     =   -1  'True
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
End
Attribute VB_Name = "frmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPersonal As ADODB.Recordset
Private Sub cmdActualizar_Click()
Dim Sql As String


Set rsPersonal = New ADODB.Recordset
rsPersonal.CursorLocation = adUseClient

Sql = " SELECT     IDPERSONAL, NOMBRE, APELLIDO,HORA_INGRESO_INICIO,HORA_INGRESO_FIN ,HORA_SALIDA_INICIO,HORA_SALIDA_FIN, NAVES, ADMINISTRATIVO, USUARIOSYS , ACTIVO"
Sql = Sql & " From Personal"
Sql = Sql & "  ORDER BY IDPERSONAL"

rsPersonal.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
Set grdPersonal.DataSource = rsPersonal.DataSource




End Sub

Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdPersonal
End Sub

Private Sub Command1_Click()
    Dim Sql As String
    Sql = " SELECT IDPERSONAL , APELLIDO, NOMBRE "
    Sql = Sql & " FROM   PERSONAL "
    frmReportes.ImprimirReporte PasoReportes + "rptPertsonal.rpt", Sql, True
End Sub

Private Sub Command3_Click()
Dim fecha As String
Dim I As Integer
Dim Sql As String

fecha = "01/01/2014"
For I = 1 To 3000


Sql = " Insert Into Dia("
Sql = Sql & " TIPO_DIA"
Sql = Sql & ",Dia"
Sql = Sql & ",NOMBRE_DIA"
Sql = Sql & ")"
Sql = Sql & " VALUES  ("
Sql = Sql & "'LABORABLE'"
Sql = Sql & ",'" & fecha & "'"
Sql = Sql & ",'" & Format(fecha, "dddd") & "'"
Sql = Sql & ")"
ExecutarSql Sql

fecha = DateAdd("D", 1, fecha)
Next

End Sub

Private Sub Form_Resize()
grdPersonal.Width = frmPersonal.Width - 400
grdPersonal.Height = frmPersonal.Height - 1800

End Sub
