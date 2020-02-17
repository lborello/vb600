VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIfrmInicio 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Requerimientos"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14730
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
      Align           =   3  'Align Left
      Height          =   10050
      Left            =   0
      OleObjectBlob   =   "mdiEntradaSalida.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   1995
   End
   Begin MSComctlLib.StatusBar StaInicio 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   10050
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ventana"
      Begin VB.Menu mnuCascada 
         Caption         =   "Cascada"
      End
   End
End
Attribute VB_Name = "MDIfrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
    On Error GoTo salir
    Select Case pLink.Caption
    Case "Actualizacion"
     
    
     Case "Planta"
     
     frmControlEstados.Show
frmControlEstados.SetFocus
    Case "Control de Remitos"
   
        frmControlRemito.Show
   frmControlRemito.SetFocus
    
    Case "dxIRequerimientos"
          
            frmCargarRequerimientos.SetFocus
            frmCargarRequerimientos.Show
    Case "dxIHojaRuta"
            frmHojaRuta.CargarModificacion (InputBox("Ingrese el Nº de Hoja de Ruta", "Hoja", 0))
            frmHojaRuta.SetFocus
            frmHojaRuta.Show
    Case "dxIRemitosEntrada"
    
    
            frmRemitoEntrada.Show
            frmRemitoEntrada.CargarRemitoEntrada
     
'    Case "Correo"
'        frmCorreo.Show
'        frmCorreo.SetFocus
     
     Case "RemitoDigital"
            
            frmImagenes.Show
      Case "dxIUsuario"
           
            frmUsuariosClientes.Show
      Case "Movimientos"
            frmMovimientos.SetFocus
            frmMovimientos.Show
       Case "ControlReferencia"
        If InputBox("Ingrese la clave") = "29425452" Then
            frmControlReferencia.SetFocus
            frmControlReferencia.Show
         Else
         MsgBox "Usuario Incorrecto"
        End If
       Case "Terminado"
        frmEstadoRequerimientoTerminado.SetFocus
        frmEstadoRequerimientoTerminado.Show
            
     End Select
MDIfrmInicio.Arrange 0
Exit Sub
salir:
    MsgBox Err.Description

End Sub

Private Sub MDIForm_Load()

    Inicio
    
    
    Rem  strConBasa , 0 ,1.CursorLocation = adUseClient
    Rem  strConBasa , 0 ,1.CommandTimeout = 6000
    MDIfrmInicio.AutoShowChildren = True
    BaseOracle = False
End Sub

Private Sub mnuCascada_Click()
 Me.Arrange 0
End Sub

