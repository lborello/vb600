VERSION 5.00
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmPersonal 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal"
   ClientHeight    =   945
   ClientLeft      =   2745
   ClientTop       =   2925
   ClientWidth     =   4110
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4110
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   540
      Width           =   1155
   End
End
Attribute VB_Name = "frmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Dim IDREQUERIMIENTO As Long
    On Error GoTo salir
    
    
    
    
    Dim rs As New ADODB.Recordset
    Dim sql  As String
    
    If IsNull(ctlPersonal.Valor) Then
        MsgBox "Ingrese el personal"
        Exit Sub
        
    End If
    
    
    
    
    
    sql = " SELECT     IDPERSONAL, NOMBRE, APELLIDO"
    sql = sql & " From dbo.Personal "
    sql = sql & "  Where IDPERSONAL = " & ctlPersonal.Valor
    rs.Open sql, ConActiva, 0, 1
    
    If rs.EOF Then
        MsgBox "Personal Error"
       ctlPersonal.Valor = 99
       Exit Sub
       
    End If
    
    
     
     
    
    CRequerimientos.CambioEstado ctlPersonal.Valor, False, 2, EstadoFinal, ConActiva
    IDREQUERIMIENTO = CRequerimientos.Item(1).NumeroRequerimiento
    Unload Me
    Select Case TRequerimientos.Item(CStr(IDREQUERIMIENTO)).IDTIPOREQUERIMIENTO
    Case 1, 3, 9 ' Cajas
        frmControlEstados.ImprimirRequerimientoCajas IDREQUERIMIENTO
    Case 5
         frmControlEstados.ImprimirIndiceRequerimiento IDREQUERIMIENTO
    Case 2, 4 ' Libros
         frmControlEstados.ImprimirRequerimientoLibros IDREQUERIMIENTO
    Case 10, 11 'Legajos
         frmControlEstados.ImprimirRequerimientoLegajos CStr(IDREQUERIMIENTO)
    Case 13
        frmControlEstados.ImprimirRequerimientoGeneral CRequerimientos.Item(1).NumeroRequerimiento
    Case Else
    End Select
    frmControlEstados.CargarTree
    Exit Sub
salir:
 frmControlEstados.CargarTree
MsgBox Err.Description
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ctlPersonal.TipoControl = PERSONAL
   
End Sub


