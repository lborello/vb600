VERSION 5.00
Object = "{A30B2DDF-C00F-469F-A23C-D6177608A128}#10.5#0"; "crviewer.dll"
Begin VB.Form frmReportes 
   Caption         =   "Reportes Sistema Requerimientos"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   15555
   Begin CrystalActiveXReportViewerLib105Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   10395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      _cx             =   26908
      _cy             =   18336
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ImprimirReporte(Paso As String, sql As String, Visible As Boolean, Optional Formula1 As String, Optional Formula2 As String, Optional CantidadCopias As Integer)
    Dim AdoRs As New ADODB.Recordset
    Dim a As New CRAXDDRT.Application
    Dim R As New CRAXDDRT.Report
    Dim fo As FormulaFieldDefinitions
    Dim i As Integer
    On Error GoTo salir
        Set R = a.OpenReport(Trim(Paso))
        Dim conreporte As New ADODB.Connection
        
        conreporte.Open strConBasa
        conreporte.CursorLocation = adUseClient
        conreporte.CommandTimeout = 100
        AdoRs.Open sql, conreporte, 0, 1
        R.DiscardSavedData
        
            
        For i = 1 To R.Database.Tables.Count
            R.Database.Tables(i).SetDataSource AdoRs
        Next
       
        If Formula1 <> "" Then
            R.FormulaFields.Item(1).Text = "'" & Formula1 & "'"
        End If
        If Formula2 <> "" Then
            R.FormulaFields.Item(2).Text = "'" & Formula2 & "'"
        End If
        R.Database.SetDataSource AdoRs, 3, 1
        If Visible = True Then
            Screen.MousePointer = vbHourglass
            CRViewer1.ReportSource = R
            CRViewer1.ViewReport
            Screen.MousePointer = vbDefault
            frmReportes.Show
        Else
        
        
        
   
            If CantidadCopias = 0 Then
                R.PrintOut False
             Else
                R.PrintOut False, CantidadCopias
             End If
        End If
        
        Exit Sub
salir:
        MsgBox Err.Description

End Sub

Public Sub Exportarpdf(Paso As String, sql As String, PasoDestino As String)
    Dim AdoRs As New ADODB.Recordset
    Dim a As New CRAXDDRT.Application
    Dim R As New CRAXDDRT.Report
    Dim fo As FormulaFieldDefinitions
    Dim i As Integer
    On Error GoTo salir
        Set R = a.OpenReport(Trim(Paso))
        Dim conreporte As New ADODB.Connection
        
        conreporte.Open strConBasa
        conreporte.CursorLocation = adUseClient
        conreporte.CommandTimeout = 100
        AdoRs.Open sql, conreporte, 0, 1
        R.DiscardSavedData
        
            
        For i = 1 To R.Database.Tables.Count
            R.Database.Tables(i).SetDataSource AdoRs
        Next
       
       R.ExportOptions.FormatType = crEFTPortableDocFormat
R.ExportOptions.DestinationType = crEDTDiskFile
R.ExportOptions.DiskFileName = PasoDestino
    R.Export False
    
       
        Exit Sub
salir:
        MsgBox Err.Description

End Sub


Public Sub ExportarReporte(Paso As String, sql As String, Visible As Boolean, Optional Formula1 As String, Optional Formula2 As String, Optional Archivo As String)
    Dim AdoRs As New ADODB.Recordset
    Dim a As New CRAXDDRT.Application
    Dim R As New CRAXDDRT.Report
    Dim fo As FormulaFieldDefinitions
    Dim i As Integer
    On Error GoTo salir
        Set R = a.OpenReport(Trim(Paso))
        Dim conreporte As New ADODB.Connection
        
        conreporte.Open strConBasa
        conreporte.CursorLocation = adUseClient
        conreporte.CommandTimeout = 100
        AdoRs.Open sql, conreporte, 0, 1
        R.DiscardSavedData
        For i = 1 To R.Database.Tables.Count
            R.Database.Tables(i).SetDataSource AdoRs
        Next
       
        If Formula1 <> "" Then
            R.FormulaFields.Item(1).Text = "'" & Formula1 & "'"
        End If
        If Formula2 <> "" Then
            R.FormulaFields.Item(2).Text = "'" & Formula2 & "'"
        End If
        
        
        
        R.ExportOptions.ExcelExportAllPages = True
    R.ExportOptions.DestinationType = crEDTDiskFile
    R.ExportOptions.FormatType = crEFTTabSeparatedText
    R.ExportOptions.DiskFileName = Archivo
    R.Export (False)
        
        Exit Sub
salir:
        MsgBox Err.Description

End Sub

Private Sub Form_Resize()
  CRViewer1.Width = frmReportes.Width - 200
  CRViewer1.Height = frmReportes.Height - 200
End Sub
