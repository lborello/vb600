VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiEntradaSalida 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Requerimientos (PARALELO SQL)"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8535
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6045
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
      Align           =   3  'Align Left
      Height          =   6045
      Left            =   0
      OleObjectBlob   =   "mdiEntradaSalida.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ventana"
      Begin VB.Menu mnuCascada 
         Caption         =   "Cascada"
      End
   End
End
Attribute VB_Name = "mdiEntradaSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
    Select Case pLink.Caption
    
    Case "dxIRequerimientos"
            frmCargarRequerimientos.SetFocus
            frmCargarRequerimientos.Show
    Case "dxIHojaRuta"
            frmHojaRuta.CargarModificacion (InputBox("Ingrese el Nº de Hoja de Ruta"))
            frmHojaRuta.SetFocus
            frmHojaRuta.Show
    Case "dxIRemitosEntrada"
            frmRemitoEntrada.Show
            frmRemitoEntrada.CargarRemitoEntrada
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
       
            
     End Select
mdiEntradaSalida.Arrange 0
End Sub

Private Sub MDIForm_Load()
 Set ConBasa = New ADODB.Connection
    UserName = "BASA"
    Password = "1742"
    Rem ConBasa.Open "Provider=MSDAORA.1;Password=1742;User ID=basa;Data Source=bpdc;Persist Security Info=True"
     ConBasa.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BasaSistema;Data Source=SERVER001"
    mdiEntradaSalida.AutoShowChildren = True
    BaseOracle = False
End Sub

Private Sub mnuCascada_Click()
 Me.Arrange 0
End Sub

