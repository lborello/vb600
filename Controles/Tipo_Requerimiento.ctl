VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK~1.OCX"
Begin VB.UserControl Tipo_Requerimiento 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   ScaleHeight     =   405
   ScaleWidth      =   6030
   Begin VB.ComboBox cboRazon_Social 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin MSMask.MaskEdBox mskCod_Cliente 
      Height          =   360
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Tipo_Requerimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

