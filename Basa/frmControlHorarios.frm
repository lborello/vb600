VERSION 5.00
Begin VB.Form frmControlHorarios 
   Caption         =   "Control de Horarios"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6720
      Top             =   0
   End
   Begin VB.TextBox txtSalida 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   2475
   End
   Begin VB.TextBox txtEntrada 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label lblUsuario 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1740
      Width           =   7035
   End
   Begin VB.Label lblHorario 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   7035
   End
   Begin VB.Label Label2 
      Caption         =   "Salida:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Entrada:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmControlHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
lblHorario.Caption = Now
End Sub
