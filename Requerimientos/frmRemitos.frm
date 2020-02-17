VERSION 5.00
Begin VB.Form frmRemitos 
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   795
      Left            =   1560
      TabIndex        =   4
      Top             =   2820
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Sector 
      Caption         =   "Sector"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Solicitante :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label lblRazonSocial 
      Caption         =   "Label2"
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmRemitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim A As Integer
Dim I As Integer
 A = InputBox("LLL", "LLL")


Select Case A
Case 1
MsgBox A
Case 2, 3, 4
MsgBox A
Case A To 10


MsgBox "KKKKK"
End Select

For I = 10 To A
 Beep

Next








End Sub

