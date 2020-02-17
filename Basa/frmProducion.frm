VERSION 5.00
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmProducion 
   Caption         =   "Produccion"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   8145
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   6480
      TabIndex        =   6
      Top             =   2220
      Width           =   1455
   End
   Begin VB.TextBox txtUnidades 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   660
      Width           =   1815
   End
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
   End
   Begin VB.TextBox txtDescricion 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Unidades"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Personal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmProducion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub cmdAceptar_Click()
    Dim Sql As String

    Sql = " INSERT INTO TAREAS (ID_PERSONAL, DESCRIPCION, UNIDADES, FECHA)"
    Sql = Sql & " VALUES (" & ctlPersonal.Valor & ",'" & Trim(txtDescricion.Text) & "'," & txtUnidades.Text & "," & SysDate & ")"
    ExecutarSql Sql
MsgBox "Terminado"
Unload Me


End Sub

Private Sub Form_Load()
    ctlPersonal.TipoControl = PERSONAL
End Sub
