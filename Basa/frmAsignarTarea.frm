VERSION 5.00
Begin VB.Form frmAsignarTarea 
   Caption         =   "Asignar Tarea"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   4245
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   5220
      Width           =   1200
   End
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
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   5220
      Width           =   1200
   End
   Begin VB.ListBox lstAsignacionResponsable 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   540
      Width           =   4095
   End
   Begin VB.Label lblTipoTarea 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Tarea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   120
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Tarea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmAsignarTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CargarLista(FiltroNombreCampo As String, ValorCampo As Integer, AsignarTipoTarea As TipoTarea)
 Dim rsCargarLista As ADODB.Recordset
 Set rsCargarLista = New ADODB.Recordset
 Dim sSQL As String
    sSQL = " SELECT IDPERSONAL, NOMBRE, APELLIDO, NAVES,"
    sSQL = sSQL & vbCrLf & " ADMINISTRATIVO , CHOFER, JEFEPLANTA, USUARIOSYS"
    sSQL = sSQL & vbCrLf & " FROM PERSONAL"
    sSQL = sSQL & vbCrLf & " WHERE    "
    sSQL = sSQL & vbCrLf & FiltroNombreCampo & " = " & ValorCampo
    rsCargarLista.Open sSQL, ConActiva, 0, 1
    lstAsignacionResponsable.Clear
    Set Responsables = New clsResponsables
    Do While Not rsCargarLista.EOF
         lstAsignacionResponsable.AddItem Format(rsCargarLista!idPersonal, "000") & " - " & Trim(rsCargarLista!Apellido) & " " & Trim(rsCargarLista!Nombre)
        rsCargarLista.MoveNext
    Loop
    rsCargarLista.Close
    lblTipoTarea.Caption = AsignarTipoTarea
End Sub

Private Sub cmdAceptar_Click()
 Dim i As Integer
 For i = 0 To lstAsignacionResponsable.ListCount - 1
   If lstAsignacionResponsable.Selected(i) = True Then
        Responsables.Add "P" & Mid(lstAsignacionResponsable.List(i), 1, 3), lstAsignacionResponsable.List(i), CInt(Mid(lstAsignacionResponsable.List(i), 1, 3)), "P" & Mid(lstAsignacionResponsable.List(i), 1, 3)
   End If
 Next
 frmAsignarTarea.Hide
 
End Sub

