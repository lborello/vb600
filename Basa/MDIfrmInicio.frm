VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIfrmInicio 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "BIENVENIDOS A SISTEMA BASA"
   ClientHeight    =   9165
   ClientLeft      =   1560
   ClientTop       =   810
   ClientWidth     =   15240
   Icon            =   "MDIfrmInicio.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIfrmInicio.frx":030A
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StaInicio 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8790
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
      Align           =   3  'Align Left
      Height          =   8790
      Left            =   0
      OleObjectBlob   =   "MDIfrmInicio.frx":2C5114
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "MDIfrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
   Select Case pLink.ObjectName
   Case "EnvioCorreo"
   frmEnvioCorreo.Show
   frmEnvioCorreo.SetFocus
  Case "migracion"
        frmMigracion.Show
        frmMigracion.SetFocus
   Case "dxIFacturacion"
        frmfacturacionTango.Show
        frmfacturacionTango.SetFocus
   Case "Personal"
        frmPersonal.Show
        frmPersonal.SetFocus
   Case "Andesmar"
        frmAndesmar.Show
        frmAndesmar.SetFocus
   Case "Informe Legajos"
        Dim Sql As String
 
' SELECT RAZON_SOCIAL, COD_CLIENTE, DESCRIPCION_REMITO, NRO_CAJA,ID_CLIENTE_LEGAJO
' FROM   "BasaSistema"."dbo"."V_REPORTE_LEGAJOS" "V_REPORTE_LEGAJOS
' ORDER BY "V_REPORTE_LEGAJOS"."RAZON_SOCIAL", "V_REPORTE_LEGAJOS"."NRO_CAJA"

Dim Cliente As Integer

        Cliente = InputBox("Ingrese el cliente", "0", "0")
        Sql = " SELECT * "
        Sql = Sql & " From V_LEGAJOS_DOCUMENTOS"
        Sql = Sql & " where COD_CLIENTE = " & Cliente
        Sql = Sql & " and INDICE like '" & BuscarIndice(Cliente, InputBox("Ingrese el n° documento", "", "0")) & "%'"
        ' SQL = " SELECT RAZON_SOCIAL, DESCRIPCION_REMITO, NRO_CAJA, ID_CLIENTE_LEGAJO , LEGAJOS.COD_CLIENTE"
        ' SQL = SQL & " FROM  LEGAJOS INNER JOIN CLIENTES ON LEGAJOS.COD_CLIENTE=CLIENTES.ID_CLIENTE"
        ' SQL = SQL & " where LEGAJOS.COD_CLIENTE = " & InputBox("Ingrese el cliente", "", "0")
        ' SQL = SQL & "  ORDER BY NRO_CAJA"
 
  frmReportes.ImprimirReporte PasoReportes + "rptLegajosDocumentos.rpt", Sql, True
   Case "contenedor"
   frmContenedor.Show
   frmContenedor.SetFocus
   
   Case "ControlHorario"
    frmControlHorario.Show
    frmControlHorario.SetFocus
   
   
   Case "Padron"
        frmPadron.Show
        frmPadron.SetFocus
   Case "Cajas"
        frmCajas.Show
        frmCajas.SetFocus
   Case "dxIGenerarCajas"
        frmGeneracionCajas.Show
        frmGeneracionCajas.SetFocus
   
   Case "dxICrearCliente"
        frmClienteMaestro.Show
        frmClienteMaestro.SetFocus
   Case "dxIControl"
         frmControlesVarios.Show
          frmControlesVarios.SetFocus
   Case "dxICobranzas"

   Case "dxIEnvioInfo"
         frmEnvioInfo.Show
         frmEnvioInfo.SetFocus
   Case "dxIImportarExcel"
         frmImportExcel.Show
         frmImportExcel.SetFocus
   Case "dxIEtiquetasLegajos"
        frmEtiquetasLegajos.Show
        frmEtiquetasLegajos.SetFocus
   Case "dxICambioHIstorico"
        FrmCambioPosicionHistorico.Show
        FrmCambioPosicionHistorico.SetFocus
   Case "dxIEntrada"
        frmEntrada.Show
        frmEntrada.SetFocus
   Case "dxISimulacionLectura"
        frmLecturaSimulacion.Show
        frmLecturaSimulacion.SetFocus
   Case "dxIOsepAfiliados"
        frmOsepAfiliados.Show
        frmOsepAfiliados.SetFocus
 Case "dxIConsultasDigitales"
'        frmConsultaDIgital.Show
'        frmConsultaDIgital.SetFocus
  Case "dxlOsepRecetas"
        frmRecetasOsep.Show
        frmRecetasOsep.SetFocus
  Case "dxIPrecintos"
        frmPrecintos.Show
        frmPrecintos.SetFocus
  Case "dxICambioPosicion"
        
        frmCambioPosicionFisica.Show
        frmCambioPosicionFisica.SetFocus
  Case "dxIMovimientosLegajos"
        
        FrmMovimientosLegajos.Show
        FrmMovimientosLegajos.SetFocus
  Case "dxIEcogas"
        FrmEcogas.Show
        FrmEcogas.SetFocus
  Case "dxIHorasArchivista"
        frmArchivistaPorHora.Show
        frmArchivistaPorHora.SetFocus
  Case "dxIOrdenDocumentacion"
        frmOrdenDocumentacion.Show
        frmOrdenDocumentacion.SetFocus
        Rem frmOrdenamientoDocumentacion.Show
        Rem frmOrdenamientoDocumentacion.SetFocus
  Case "dxICargarLegajos"
        frmAsignacionCajasCarga.Show
        frmAsignacionCajasCarga.SetFocus
  Case "dxIInformacionGerencial"
        frmImfomacionGerencial.Show
        frmImfomacionGerencial.SetFocus
   Case "dxIUbicacionLibros"
        frmLibrosUbicacion.Show
        frmLibrosUbicacion.SetFocus
   Case "dxIUbicacionCajas"
        frmCajasUbicacion.Show
        frmCajasUbicacion.SetFocus
  Case "CajasRotulos"
        frmImpresionRotulo.Show
        frmImpresionRotulo.SetFocus
  Case "CajasMovimientos"
        FrmMovimientosCajas.Show
        FrmMovimientosCajas.SetFocus
  Case "RequerimientosCargar"
        Rem  frmCargarRequerimientos.Show
  Case "RequerimientosControldeSalidas"
  Case "dxIBuscarLegajos"
    frmBuscarLegajos.Show
    frmBuscarLegajos.SetFocus
  Case "dxILibrosReferencias"
    frmCargarLibros.Show
    frmCargarLibros.SetFocus
  Case "LecturaColector"
    frmLecturaMemo.Show
    frmLecturaMemo.SetFocus
  Case "CajasReferencias"
    
        frmAgregarDocumentos.Show
        frmAgregarDocumentos.SetFocus
           
    Case "ReferenciasAnterior"
     frmCajasReferencias.Show
            frmCajasReferencias.SetFocus
            
  Case "Buscar"
    frmBuscarGenerico.Show
     frmBuscarGenerico.SetFocus
    
  Case "LibrosReferencias"
   Rem  frmLibrosReferencias.Show
  Case "dxISolicitante"
   
  Case "ReUbicacion"
     frmPosicionamiento.Show
     frmPosicionamiento.SetFocus
  Case "dxIPermisos"
    frmUsuariosClientes.Show
    frmUsuariosClientes.SetFocus
  
  Case "dxIProducion"
    
     If MDIfrmInicio.StaInicio.Panels(2) = 47 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Then
        frmProducion.Show
        frmProducion.SetFocus
     End If
    
    
  Case "dxIIndices"
    
    If MDIfrmInicio.StaInicio.Panels(2) = 17 Or MDIfrmInicio.StaInicio.Panels(2) = 47 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Or MDIfrmInicio.StaInicio.Panels(2) = 31 Or MDIfrmInicio.StaInicio.Panels(2) = 82 Or MDIfrmInicio.StaInicio.Panels(2) = 19 Or MDIfrmInicio.StaInicio.Panels(2) = 69 Then
        frmIndices.Show
        frmIndices.SetFocus
    Else
        MsgBox "Solicite la creación de índices a administración"
    End If
  
  
  Case "dxIBloques"
    frmBloque.Show
    frmBloque.SetFocus
'  Case "dxIRequerimientos"
'    frmRequerimientos.Show
'    frmRequerimientos.SetFocus
  Case "dxIUnificarLegajos"
    frmUnificarLegajos.Show
    frmUnificarLegajos.SetFocus
   Case "dxIManejoArchivos"
    frmManejoArchivo.Show
    frmManejoArchivo.SetFocus
   Case "dxIndexarImagenes"
    frmIndexarImganenes.Show
    frmIndexarImganenes.SetFocus
  Case "Tomador"
    frmTomador.Show
    frmTomador.SetFocus
 
 
 End Select
  Me.Arrange 0
  Debug.Print pLink.ObjectName
  
End Sub

Private Sub MDIForm_Load()
    inicio

End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
 MsgBox X & " y" & Y
 End If
End Sub

Private Sub MDIForm_Terminate()

'    If Not oWordL Is Nothing Then
'       MousePointer = 11
'       If Not oTmpDocL Is Nothing Then
'            oTmpDocL.Close
'       End If
'       oWordL.Quit
'       Set oWordL = Nothing
'       MousePointer = 0
'    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
' If Not oWordL Is Nothing Then
'    MousePointer = 11
'    If Not oTmpDocL Is Nothing Then
'            oTmpDocL.Close
'    End If
'
'   If Not IsEmpty(oWordL) Then
'     oWordL.Quit
'    End If
'    Set oWordL = Nothing
'    MousePointer = 0
' End If
End Sub

Private Sub mnuventana_Click()
    Me.Arrange vbCascade
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
  Case "Cajas"
  Case "Memo"
    frmLecturaMemo.Show
  Case "Requerimientos"
   
  End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Text
Case "Imprimir"
    frmImpresionRotulo.Show
Case "Ubicacion"
    frmCajasUbicacion.Show
Case "Registracion"
    Rem frmMovimientosCajas.Show
Case "Movimientos"
    FrmMovimientosCajas.Show
Case "Control de Salida"
   
Case "Posicionamiento"
     frmPosicionamiento.Show
End Select
End Sub

