VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportExcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Excel"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
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
   ScaleHeight     =   6720
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   915
      Left            =   60
      TabIndex        =   16
      Top             =   3420
      Width           =   6255
      Begin VB.CommandButton cmdOsepRecetas 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   3900
         TabIndex        =   19
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtQuincenaMesAño 
         Height          =   375
         Left            =   2340
         TabIndex        =   17
         Text            =   "01102014"
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Quincena Mes Año "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cambio de Indices"
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   4500
      Width           =   6255
      Begin VB.CommandButton cmdRemitosNuevos 
         Caption         =   "Remitos Nuevos y requerimientos"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton cmdRecibos 
         Caption         =   "Recibos"
         Height          =   435
         Left            =   4620
         TabIndex        =   15
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdRequerimiento 
         Caption         =   "Requerimientos"
         Height          =   435
         Left            =   3120
         TabIndex        =   11
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemitos 
         Caption         =   "Remitos"
         Height          =   435
         Left            =   1620
         TabIndex        =   10
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdPercistenciaDocumetos 
         Caption         =   "Proceso"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Carga de referencia"
      Height          =   2655
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.CheckBox chkBorrarReferencia 
         Caption         =   "Borrar referencia  Administrada por el cliente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1500
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CommandButton cmdControlLegajos 
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   300
         TabIndex        =   13
         Top             =   1920
         Width           =   2160
      End
      Begin VB.CommandButton cmdPlanillaManual 
         Caption         =   "Planilla Manual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3540
         TabIndex        =   12
         Top             =   1920
         Width           =   2160
      End
      Begin VB.CheckBox chkCambio_Referencia 
         Caption         =   "Cambio de Referencia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   780
         Width           =   2115
      End
      Begin VB.CheckBox chkNoControlasEstado 
         Caption         =   "No controlar Estado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   420
         Width           =   2055
      End
      Begin VB.CheckBox chkNoControlarRangos 
         Caption         =   "No controlar rangos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkControlCompleto 
         Caption         =   "Control Completo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   1755
      End
      Begin VB.CheckBox chkNoControlRefCargada 
         Caption         =   "Controlar si la referencia esta cargada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Value           =   1  'Checked
         Width           =   3855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblCantidadRegistros 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Cantidad de registros:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "frmImportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConSql As ADODB.Connection
Dim PasoServer As String
Dim PasoServerImagenSql As String
Dim ControlIndiceNumero As Boolean
Dim ControlIndiceFecha As Boolean
Dim ControlIndiceLetra As Boolean
Dim ControlIndiceTipo As String
Dim clienteCajasChicas As Long

Public Function TraerIndice(Documento As Integer, Cliente As Integer, Optional ERROR As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        Dim Sql As String
        Sql = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, "
        Sql = Sql & vbCrLf & " Descripcion , Fecha, NUMERO, LETRA, TIPO_INDICE"
        Sql = Sql & vbCrLf & "  From INDICES"
        Sql = Sql & vbCrLf & "  WHERE COD_CLIENTE =" & Cliente & " AND  ID_CODIGO_DOCUMENTO = " & Documento
        rs.Open Sql, ConActiva, 0, 1
        ControlIndiceFecha = False
        ControlIndiceLetra = False
        ControlIndiceNumero = False
        ControlIndiceTipo = ""
        ERROR = ""
        If rs.EOF Then
            TraerIndice = ""
            ERROR = "No existe el indice"
        Else
            TraerIndice = rs!Indice
            If Not IsNull(rs!fecha) Then
                ControlIndiceFecha = True
            End If
            If Not IsNull(rs!NUMERO) Then
                ControlIndiceNumero = True
            End If
            If Not IsNull(rs!lETRA) Then
                ControlIndiceLetra = True
            End If
            If Not (Trim(rs!Tipo_Indice) = "Documento" Or Trim(rs!Tipo_Indice) = "Legajo") And Cliente = 4 Then
                TraerIndice = ""
                ERROR = "El tipo de indice no es el correcto " & rs!Tipo_Indice
            End If
        End If
End Function

Public Function Control_Estado_Caja(Cliente As Integer, NRO_CAJA As Long) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT ESTADO "
        Sql = Sql & " From CONTENEDOR "
        Sql = Sql & " Where COD_CLIENTE = " & Cliente
        Sql = Sql & " And NRO_CAJA = " & NRO_CAJA
    rs.Open Sql, ConActiva, 0, 1
    Control_Estado_Caja = ""
    If rs.EOF Then
        Control_Estado_Caja = "NO existe la Caja"
    Else
        If rs!estado <> 2 Then
           If rs!estado <> 3 Then
            Control_Estado_Caja = "Caja estado: " & rs!estado
           End If
        End If
    End If
End Function

Public Function Control_Referencia_Caja(Cliente As Integer, NRO_CAJA As Long) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT * "
        Sql = Sql & "  From REFERENCIAS "
        Sql = Sql & "  Where NRO_CAJA = " & NRO_CAJA
        Sql = Sql & "  And COD_CLIENTE = " & Cliente
    rs.Open Sql, ConActiva, 0, 1
    Control_Referencia_Caja = ""
    If rs.EOF Then
        Control_Referencia_Caja = "Ref: "
    End If
End Function

Public Function Traer_Indice_Anterior(COD_ID_REFERENCIA As Long) As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
Sql = " SELECT INDICE From REFERENCIAS Where COD_ID_REFERENCIA = " & COD_ID_REFERENCIA

    rs.Open Sql, ConActiva, 0, 1
    If rs.EOF Then
        Traer_Indice_Anterior = ""
    Else
        Traer_Indice_Anterior = rs!Indice
    End If

End Function

Private Sub cmdCargaLegajos_Click()
On Error GoTo salir
 Dim Paso As String
            CommonDialog1.FileName = "\\Ser-cudea\documentos\PAOLA - VARIOS\Supervielle\*.XLS"
            CommonDialog1.ShowOpen

    If CommonDialog1.CancelError <> False Then
        Paso = CommonDialog1.FileName
        If Control_Excel_Legajos(Paso, False) = False Then
            MsgBox "Existe un errores", vbCritical, "Control Carga"
        Else
            Percistencia_legajos_supervielle Paso, True
        End If
    End If
salir:
    
End Sub

Private Sub cmdControl_Cliente_Click()
    Dim Paso As String
    On Error GoTo salir:

         CommonDialog1.FileName = "c:\*.xls"

        CommonDialog1.ShowOpen
        Paso = CommonDialog1.FileName

        If chkCambio_Referencia.value = 1 Then
          BorrarCajasPlanillas (Paso)
        End If


        If Control_Excel_Cliente(Paso, True) Then
            MsgBox "Existe un errores", vbCritical
        Else
            MsgBox "NO existe errores", vbInformation
        End If
salir:
End Sub


Private Sub Actualizar_Re()
'Dim conexel As New ADODB.Connection
'    Dim rsExel As New ADODB.Recordset
'    Dim Caja As Long
'    Dim Cliente As Integer
'    Dim SQL As String
'    Dim ID As Long
'    Dim Indice As String
'    Dim INDICE_ANTERIOR2 As String
'    Dim strCodigo As String
'    Dim rsCodigo As ADODB.Recordset
'    Cliente = txtCliente.Text
'Dim R As Integer
'   On Error GoTo SALIR
'   conexel.Open "Provider=MSDASQL.1;Persist Security Info=False;Mode=ReadWrite;Extended Properties=DBQ=" & CommonDialog1.FileName & ";DefaultDir=F:\Público\1- Osep\Notas y Cartas;Driver={Microsoft Excel Driver (*.xls)};DriverId=790;FIL=excel 5.0;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;ReadOnly=1;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
'   rsExel.Open "SELECT * FROM `Referencias$`", conexel
'    R = 0
'    CONBASA.BeginTrans
'    Do While Not rsExel.EOF And R < 1000
'      R = R + 1
'
'
'       If Not IsNull(rsExel!F12) Then
'            If IsNumeric(Replace(rsExel!F12, "'", "")) Then
'                'Indice
'                ID = Replace(rsExel!F12, "'", "")
'                If IsNumeric(rsExel!F1) Then 'Documento
'                     If TraerIndice(rsExel!F1, Cliente) = "" Then
'                        Indice = ""
'                     Else
'                        Indice = "'" & TraerIndice(rsExel!F1, Cliente) & "'"
'                     End If
'                 Else
'                     Clipboard.Clear
'                     Clipboard.SetText ID
'                     MsgBox "Error en el documento ID:" & ID & vbCrLf & "El error esta copiado en el Clipboard"
'                     Indice = ""
'
'
'                End If
'                INDICE_ANTERIOR2 = INDICE_ANTERIOR(ID)
'                If Indice <> "" And INDICE_ANTERIOR2 <> Indice Then
'                        SQL = "   Update REFERENCIAS"
'                        SQL = SQL & vbCrLf & " SET INDICE = " & Indice & ", PASOARCHIVO = '" & CommonDialog1.FileTitle & "', INDICE_ANTERIOR = '" & INDICE_ANTERIOR2 & "',"
'                        SQL = SQL & vbCrLf & " FECHA_MODIFICACION = TO_DATE('" & Format(Now, "DD/MM/YYYY") & "', 'DD/MM/YYYY'),"
'                        SQL = SQL & vbCrLf & " USUARIO_MODIFICACION = 'S_Cambio_indice'"
'                        SQL = SQL & vbCrLf & " Where COD_ID_REFERENCIA = " & ID
'                        SQL = SQL & vbCrLf & " And COD_CLIENTE = " & Cliente
'                        ExecutarSql SQL
'                End If
'
'             End If
'     End If
'       rsExel.MoveNext
'  Loop
'
' CONBASA.CommitTrans
' MsgBox "Proseso Terminado"
' Exit Sub
'SALIR:
'     MsgBox " Error en la operacion"
'     CONBASA.RollbackTrans
End Sub


'Private Function Control_Excel_Carga(Paso As String, ModificarNombre As Boolean) As Boolean
'    Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'    Dim B_Error As Boolean
'    Dim Error As String
'    Dim Caja As Long
'    Dim Cliente, i As Integer
'    Dim SQL, nombreArchivoOrigen, Indice, msgError, ControlEstado, Descripcion, FECHA_DESDE, FECHA_HASTA, LETRA_DESDE, LETRA_HASTA As String
'    Dim NRO_DESDE, NRO_HASTA, ID_SQL As Long
'    Dim strCodigo, ControlReferencia As String
'    Dim CantDias As Long
'    Dim ContadorFilasBlanco As Integer
'    Dim C_Error, C_Doc, C_Caja, C_Descripcion, C_Fecha_desde, C_Fecha_hasta, C_Desde, C_Hasta As Integer
'    Dim C As Integer
'    Dim Control_Columnas As Boolean
'
'    B_Error = False
'
'
'
'    Dim s_ConSql As String
'    Dim s_RsSql As String
'    s_ConSql = strConBasa , 0 ,1
'    Set ConSql = New ADODB.Connection
'    ConSql.Open s_ConSql
'    Dim RsSql As ADODB.Recordset
'
'Dim R As Excel.Range
'Dim h As Excel.Hyperlinks
'ContadorFilasBlanco = 0
'
''abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Open(Paso)
'    Set hojaEx = libroEx.Worksheets.Item(1)
'
'    If UCase(hojaEx.Name) <> "REFERENCIAS" Then
'           Control_Excel_Carga = True
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'             MsgBox "El nombre de la planilla no es el correcto", vbInformation
'        Exit Function
'    End If
'
'    ' Control de columnas
'
'    With hojaEx
'    For C = 1 To 10
'        Control_Columnas = False
'        If UCase(.Cells(1, C)) = "ERROR" Then
'            C_Error = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "PROCESO" Then
'            C_Doc = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "CAJA" Then
'            C_Caja = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "DESCRIPCION" Or UCase(.Cells(1, C)) = "DESCRIPCIÓN" Then
'            C_Descripcion = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "FECHA DESDE" Then
'            C_Fecha_desde = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "FECHA HASTA" Then
'            C_Fecha_hasta = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "DESDE" Then
'            C_Desde = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "HASTA" Then
'            C_Hasta = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) <> "" And Control_Columnas = False Then
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            MsgBox "Error en nombre de columna " & Cells(1, C)
'            Exit Function
'        End If
'    Next
'
'
'
'   For i = 2 To 3000
'       'Control de fin de rows
'        If Trim(hojaEx.Cells(i, C_Doc)) = "" And Trim(.Cells(i, C_Caja)) = "" Then
'            ContadorFilasBlanco = ContadorFilasBlanco + 1
'            If ContadorFilasBlanco > 15 Then
'                Exit For
'            End If
'            GoTo Saltar
'        End If
'        'Control de cliente
'
'        If IsNull(ctlClientes.Valor) Then
'            msgError = "Error en cliente"
'            B_Error = True
'            Exit Function
'        End If
'
'
'
'        'Control de caja
'         If ControlCaja(hojaEx.Cells(i, C_Caja), ctlClientes.Valor) <> "" Then
'            msgError = ControlCaja(hojaEx.Cells(i, C_Caja), ctlClientes.Valor)
'         End If
'
'        'Control de documento
'           If ControlDocumento(ctlClientes.Valor, hojaEx.Cells(i, C_Doc)) <> "" Then
'                msgError = ControlDocumento(ctlClientes.Valor, hojaEx.Cells(i, C_Doc))
'            End If
'
'        'Control de Fecha_desde
'        If ControlFechas(hojaEx.Cells(i, C_Fecha_desde), hojaEx.Cells(i, C_Fecha_hasta)) <> "" Then
'            msgError = ControlFechas(hojaEx.Cells(i, C_Fecha_desde), hojaEx.Cells(i, C_Fecha_hasta))
'        End If
'
'        'Control Numero Desde
'         If ControlDesdeHasta(hojaEx.Cells(i, C_Desde), hojaEx.Cells(i, C_Hasta)) <> "" Then
'            msgError = ControlDesdeHasta(hojaEx.Cells(i, C_Desde), hojaEx.Cells(i, C_Hasta))
'         End If
'
'         If Trim(hojaEx.Cells(i, C_Desde)) = "" And Trim(hojaEx.Cells(i, C_Fecha_desde)) = "" Then
'               msgError = "Ingrese indice secundario"
'         End If
'         If msgError <> "" Then
'            hojaEx.Cells(i, C_Error) = msgError
'            B_Error = True
'        End If
'        msgError = ""
'        lblTitulo.Caption = "Control Registros :"
'        lblTitulo.Refresh
'        lblCantidadRegistros.Caption = i
'        lblCantidadRegistros.Refresh
'
'Saltar:
'
'
'   Next
'
'
'
'
'   Control_Excel_Carga = B_Error
'End With
'
'
''For i = 1 To 3000
''
''
''  ID_referencia = InsertarReferencias(CInt(ctlClientes.Valor), .Cells(i, C_Caja), CStr(Indice), CStr(Descripcion) _
''                , CStr(FECHA_DESDE), CStr(FECHA_HASTA), CStr(NRO_DESDE), CStr(NRO_HASTA), CStr(LETRA_DESDE), CStr(LETRA_HASTA) _
''                , "Null", CStr(s_NombreArchivo), 0, "Null", 99)
''Next
''
'
'libroEx.Save
'libroEx.Close
'ApExcel.Quit
'Set hojaEx = Nothing
'Set libroEx = Nothing
'Set ApExcel = Nothing
''    If ModificarNombre Then
''        If B_Error = False Then
''            FileSystem.FileCopy Paso, Mid(Paso, 1, Len(Paso) - 4) & " Imagen_OK " & " .xls"
''            FileSystem.Kill Paso
''        End If
''    End If
'End Function
'
Private Function Control_Excel_Referencias_Cargadas(Paso As String, ModificarNombre As Boolean) As Boolean
    
    
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet
    
    Dim B_Error As Boolean
    Dim Caja As Long
    Dim Cliente As Integer
    Dim msgError As String
    Dim ErrorEstado As String
    Dim ErrorRef As String
    
    Dim ID_Sql_Imagen As Long
    Dim ID_SQL As Long
    ID_Sql_Imagen = 0
    
    Dim i As Integer
    
    
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Open(Paso)
    Set hojaEx = libroEx.Worksheets.Item(1)
        
    
    Dim C As Integer
 
    
    

    If hojaEx.Name <> "Referencias_Cargadas" Then
            MsgBox "El nombre de la planilla no es el correcto", vbInformation
            libroEx.Close
            ApExcel.Quit
            Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
        Exit Function
    End If
    B_Error = False

With hojaEx
    Cells(1, 46) = "ID_Imagen"
    Cells(1, 47) = "ID_Oracle"
    For i = 2 To 10000
       'Control de fin de rows
       If Trim(Cells(i, 3)) = "" And Trim(Cells(i, 4)) = "" Then
            Exit For
       End If
                If Cells(i, 3) = "" And Not IsNumeric(Cells(i, 3)) Then
                     msgError = "Error en cliente"
                  Else
                     'Control Estado
'                     ErrorEstado = ""
'                     For C = 4 To 20
'                         If hojaEx.Cells(I, C) <> "" Then
'                             If IsNumeric(hojaEx.Cells(I, C)) Then
'                                If Control_Estado_Caja(Cells(I, 3), hojaEx.Cells(I, C)) <> "" Then
'                                     ErrorEstado = ErrorEstado + " Estado:" & hojaEx.Cells(I, C)
'                                  B_Error = True
'                                End If
'                             Else
'                                 ErrorEstado = ErrorEstado + " Num:" & hojaEx.Cells(I, C)
'                                 B_Error = True
'                             End If
'                        End If
'                     Next
                 
                 'Control Referencia
                     ErrorRef = ""
                     For C = 4 To 20
                     If hojaEx.Cells(i, C) <> "" Then
                         If IsNumeric(hojaEx.Cells(i, C)) Then
                            If Control_Referencia_Caja(Cells(i, 3), hojaEx.Cells(i, C)) <> "" Then
                                 ErrorRef = ErrorRef + " Ref.:" & hojaEx.Cells(i, C)
                                  B_Error = True
                            End If
                         End If
                        End If
                     Next
                  End If
                 
                 
                 ID_Sql_Imagen = 0
                 ID_Sql_Imagen = ValidarImagenSql("\" & Trim(.Cells(i, 2)))
                 Rem ID_Sql_Imagen = 60000
                 
                 If ID_Sql_Imagen <> 0 Then
                     ID_SQL = ID_Sql_Imagen
                     .Cells(i, 46) = ID_Sql_Imagen
                     .Cells(i, 46).Hyperlinks.Add .Cells(i, 46), PasoServerImagenSql & ID_Sql_Imagen & ".TIF"
                 Else
                     ID_SQL = 0
                     msgError = "Error Imagen"
                     B_Error = True
                 End If
                If ErrorEstado <> "" Then
                     msgError = ErrorEstado
                End If
                If ErrorRef <> "" Then
                     If msgError <> "" Then
                         msgError = msgError + ErrorRef
                     Else
                          msgError = ErrorRef
                     End If
                 End If
                 
                 If msgError <> "" Then
                     .Cells(i, 1).Hyperlinks.Add .Cells(i, 1), PasoServer & Trim(.Cells(i, 2))
                     .Cells(i, 1) = msgError
                 Else
                  If .Cells(i, 1) <> 1 Then
                     .Cells(i, 1) = ""
                  End If
                 End If
                 msgError = ""
                 lblTitulo.Caption = "Control Registros :"
                 lblTitulo.Refresh
                 lblCantidadRegistros.Caption = i
                 lblCantidadRegistros.Refresh
       Next
   Control_Excel_Referencias_Cargadas = B_Error
End With

libroEx.Save
libroEx.Close
ApExcel.Quit
Set hojaEx = Nothing
Set libroEx = Nothing
Set ApExcel = Nothing
    If ModificarNombre Then
        If B_Error = False Then
            FileSystem.FileCopy Paso, Mid(Paso, 1, Len(Paso) - 4) & " Control_OK " & " .xls"
            FileSystem.Kill Paso
        End If
    End If
End Function
Private Function Control_Excel_Legajos(Paso As String, ModificarNombre As Boolean) As Boolean
'    Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'
'    Control_Excel_Legajos = True
'
'
'
'    Dim C_Error As Integer
'    Dim C_Documento As Integer
'    Dim C_Caja As Integer
'    Dim C_Etiqueta As Integer
'    Dim C_Tipo_Documento As Integer
'    Dim C_Nro_Documento As Integer
'    Dim C_Apellido_Nombre As Integer
'    Dim C_Fecha_Carga As Integer
'    Dim C_Personal_Carga As Integer
'    Dim C_Descripcion As Integer
'    Dim C_Fecha_desde As Integer
'    Dim C_Fecha_hasta As Integer
'
'    Dim i As Long
'
'
'
'
'    C_Error = 1
'    C_Documento = 3
'    C_Caja = 2
'    C_Etiqueta = 4
'    C_Tipo_Documento = 5
'    C_Nro_Documento = 6
'    C_Apellido_Nombre = 7
'    C_Descripcion = 8
'    C_Fecha_Carga = 9
'    C_Personal_Carga = 10
'    C_Fecha_desde = 11
'    C_Fecha_hasta = 12
'
'
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Open(Paso)
'    Set hojaEx = libroEx.Worksheets.Item(1)
'
'
'    Dim C As Integer
'
'
'
'
'    If UCase(hojaEx.Name) <> "PLANILLA PARA ENVIO POR CORREO" Then
'        MsgBox "El nombre de la planilla no es el correcto", vbInformation
'        libroEx.Close
'        ApExcel.Quit
'        Set hojaEx = Nothing
'        Set libroEx = Nothing
'        Set ApExcel = Nothing
'        Exit Function
'    End If
'
'
'Dim VcontrolEtiqueta As String
'
'With hojaEx
'
'    For i = 6 To InputBox("Ingrese la cantidad de registros") + 50
'    .Cells(i, C_Error) = ""
'      'Control de fin de rows
'       If .Cells(i, C_Documento) = "" And .Cells(i, C_Caja) = "" Then
'            .Cells(i, C_Error) = "No se registro"
'       Else
'            Rem Documento
'            If IsNumeric(.Cells(i, C_Documento)) Then
'                  If VerificarDocumentoLegajo(ctlClientes.Valor, .Cells(i, C_Documento), "", 0) = False Then
'                   .Cells(i, C_Error) = "El Nro documento no es un Legajo"
'                    Control_Excel_Legajos = False
'                  End If
'
'            Else
'               .Cells(i, C_Error) = "El Nro documento no es un numero"
'               Control_Excel_Legajos = False
'            End If
'
'
'        Rem caja
'         If ControlCaja(.Cells(i, C_Caja), ctlClientes.Valor) <> "" Then
'
'                .Cells(i, C_Error) = ControlCaja(.Cells(i, C_Caja), ctlClientes.Valor)
'                Control_Excel_Legajos = False
'         End If
'
'
'        If IsNumeric(.Cells(i, C_Nro_Documento)) And IsNumeric(.Cells(i, C_Etiqueta)) Then
'
'
'            Rem etiqueta
'            Rem PARASUBIR
'            VcontrolEtiqueta = controlEtiqueta(.Cells(i, C_Etiqueta), CStr(.Cells(i, C_Nro_Documento)))
'            If VcontrolEtiqueta <> "" Then
'               .Cells(i, C_Error) = VcontrolEtiqueta
'               Control_Excel_Legajos = False
'
'            End If
'        Else
'            If Trim(.Cells(i, C_Nro_Documento)) <> "NO TIENE" Then
'                .Cells(i, C_Error) = "El numero de documento o etiqueta no es valido"
'                Control_Excel_Legajos = False
'            End If
'        End If
'
'       If Not IsDate(.Cells(i, C_Fecha_Carga)) Then
'            .Cells(i, C_Error) = "FECHA DE CARGA"
'            Control_Excel_Legajos = False
'       End If
'
'
'       Rem PERSONAL
'       If Not IsNumeric(.Cells(i, C_Personal_Carga)) Then
'            .Cells(i, C_Error) = "PERSONAL CARGA"
'            Control_Excel_Legajos = False
'       Else
'            If ControlPersonal(.Cells(i, C_Personal_Carga)) = False Then
'                .Cells(i, C_Error) = "PERSONAL CARGA"
'                Control_Excel_Legajos = False
'            End If
'       End If
'
'
'
'                        If .Cells(i, C_Fecha_desde) <> "" Then
'                               If Not IsDate(.Cells(i, C_Fecha_desde)) Then
'                                   Cells(i, C_Error) = "FECHA_DESDE"
'                                        Control_Excel_Legajos = False
'                               Else
'                                   If Len(.Cells(i, C_Fecha_desde)) = 8 Or Len(.Cells(i, C_Fecha_desde)) = 10 Then
'                                        If CDate(.Cells(i, C_Fecha_desde)) < "01/01/1940" Then
'                                           .Cells(i, C_Error) = "FECHA_DESDE"
'                                             Control_Excel_Legajos = False
'                                         End If
'                                    Else
'                                        .Cells(i, C_Error) = "FECHA_DESDE"
'                                        Control_Excel_Legajos = False
'                                    End If
'                               End If
'                            End If
'
'                            'Control Fecha Hasta
'                            If .Cells(i, C_Fecha_hasta) <> "" Then
'                               If Not IsDate(.Cells(i, C_Fecha_hasta)) Then
'                                    .Cells(i, C_Error) = "FECHA_HASTA"
'                                     Control_Excel_Legajos = False
'                                End If
'                            Else
'                                If .Cells(i, C_Fecha_hasta) <> "" Then
'                                    .Cells(i, C_Error) = "FECHA_HASTA"
'                                     Control_Excel_Legajos = False
'                                End If
'                            End If
'                            'Control  Desde
'
'
'
'
'End If
'
'           If Control_Excel_Legajos = False Then
'           Rem  MsgBox "sss"
'           End If
'
'            Debug.Print i & "  " & .Cells(i, C_Error)
'    Next
'
'End With
'
'libroEx.Save
'libroEx.Close
'ApExcel.Quit
'Set hojaEx = Nothing
'Set libroEx = Nothing
'Set ApExcel = Nothing
'    If ModificarNombre Then
'        If Control_Excel_Legajos = True Then
'            FileSystem.FileCopy Paso, Mid(Paso, 1, Len(Paso) - 4) & " Control_OK " & " .xls"
'            FileSystem.Kill Paso
'        End If
'    End If
End Function

Private Function Control_Excel_Legajos_Supervielle(Paso As String, ModificarNombre As Boolean) As Boolean
'    Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'    Dim Sql As String
'    Dim i As Integer
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Open(Paso)
'    Set hojaEx = libroEx.Worksheets.Item(1)
'
'    If hojaEx.Name <> "Carga_Legajos" Then
'        MsgBox "El nombre de la planilla no es el correcto", vbInformation
'        libroEx.Close
'        ApExcel.Quit
'        Set hojaEx = Nothing
'        Set libroEx = Nothing
'        Set ApExcel = Nothing
'        Exit Function
'    End If
'
'
'With hojaEx
'    Cells(1, 19) = "ID_Imagen"
'    Cells(1, 20) = "ID_Oracle"
'    For i = 2 To 10000
'       'Control de fin de rows
'       If Cells(i, 2) = "" And Cells(i, 3) = "" Then
'            Exit For
'       End If
'
'       If IsNumeric(.Cells(i, 4)) Then
'          Indice = TraerIndice(.Cells(i, 4), .Cells(i, 2))
'            If Indice = "" Then
'                msgError = "DOC."
'                B_Error = True
'            End If
'        Else
'            msgError = "DOC."
'            B_Error = True
'        End If
'
'        ID_Sql_Imagen = 0
'        ID_Sql_Imagen = ValidarImagenSql("\" & Trim(.Cells(i, 10)) & "\" & Trim(.Cells(i, 11)) & ".tif")
'         If ID_Sql_Imagen <> 0 Then
'            ID_SQL = ID_Sql_Imagen
'            .Cells(i, 19) = ID_Sql_Imagen
'            .Cells(i, 19).Hyperlinks.Add .Cells(i, 46), PasoServerImagenSql & ID_Sql_Imagen & ".TIF"
'        Else
'            ID_SQL = 0
'            msgError = "Error Imagen"
'            B_Error = True
'        End If
'
'
'       If ErrorEstado <> "" Then
'            msgError = ErrorEstado
'       End If
'       If ErrorRef <> "" Then
'            If msgError <> "" Then
'                msgError = msgError + ErrorRef
'            Else
'                 msgError = ErrorRef
'            End If
'        End If
'
'        If msgError <> "" Then
'            .Cells(i, 3).Hyperlinks.Add .Cells(i, 3), PasoServer & Trim(.Cells(i, 10)) & "\" & Trim(.Cells(i, 11)) & ".Tif"
'            .Cells(i, 1).Hyperlinks.Add .Cells(i, 1), PasoServer & Trim(.Cells(i, 10)) & "\" & Trim(.Cells(i, 11)) & ".Tif"
'            .Cells(i, 1) = msgError
'        Else
'        .Cells(i, 3).Hyperlinks.Add .Cells(i, 3), PasoServer & Trim(.Cells(i, 10)) & "\" & Trim(.Cells(i, 11)) & ".Tif"
'         If .Cells(i, 1) <> 1 Then
'            .Cells(i, 1) = ""
'         End If
'        End If
'        msgError = ""
'        lblTitulo.Caption = "Control Registros :"
'        lblTitulo.Refresh
'        lblCantidadRegistros.Caption = i
'        lblCantidadRegistros.Refresh
'    Next
'
'End With
'
'libroEx.Save
'libroEx.Close
'ApExcel.Quit
'Set hojaEx = Nothing
'Set libroEx = Nothing
'Set ApExcel = Nothing
'    If ModificarNombre Then
'        If Control_Excel_Legajos_Supervielle = False Then
'            FileSystem.FileCopy Paso, Mid(Paso, 1, Len(Paso) - 4) & " Control_OK " & " .xls"
'            FileSystem.Kill Paso
'        End If
'    End If
End Function


Private Function Control_Excel_Cliente(Paso As String, ModificarNombre As Boolean) As Boolean
        Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim B_Error As Boolean
        Dim i As Integer
        Dim msgError  As String
        Dim UltimaFila As Integer
        Dim ErrorDoc As String
        Dim ERRORCAJA As String
        Dim ValidarIndiceSecundario As Boolean
        Dim fkcliente As Integer


       On Error GoTo salir
        'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(Paso)
        Set hojaEx = libroEx.Worksheets.Item(1)


        'Control de Nombre de planilla
        If hojaEx.Name <> "Planilla para envio por correo" Then
            MsgBox "El nombre de la planilla no es el correcto" & vbCrLf & " El Nombre correcto es:" & "Planilla para envio por correo", vbInformation
            Control_Excel_Cliente = True
            libroEx.Close
            ApExcel.Quit
            Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
            Exit Function
        End If

        'Control de formato
        If Not (Mid(Cells(4, 3), 1, 18) <> "Nombre:" Or Mid(Cells(4, 3), 1, 18) <> "Nombre y Apellido:") Then
            MsgBox "Error en el formato", vbInformation
            Control_Excel_Cliente = True
            libroEx.Close
            ApExcel.Quit
            Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
            Exit Function
        End If

        'Control de Cliente
'        If IsNull(fkcliente) Then
'            MsgBox "Ingrese el cliente", vbCritical
'            Control_Excel_Cliente = True
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            Exit Function
'        End If
    
    fkcliente = 4
        If Cells(6, 2) <> "Caja" Then
            MsgBox "Formato de planilla incorrecto", vbCritical
            Control_Excel_Cliente = True
            libroEx.Close
            ApExcel.Quit
            Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
            Exit Function
        End If

        hojaEx.Unprotect 21877471
        'Iniciaclizacion de Bandera
        B_Error = False
        Dim FinPlanilla As Boolean
        Dim ContarBlanco As Integer
        FinPlanilla = False
        ContarBlanco = 0

        With hojaEx
        For i = 7 To 6000
               'Control de fin de rows
                If Cells(i, 2) = "" Or Cells(i, 3) = "" Then
                    FinPlanilla = True
                    ContarBlanco = ContarBlanco + 1
                    If ContarBlanco > 20 Then
                        Exit For
                    End If
                 Else
                       ValidarIndiceSecundario = False
                       If FinPlanilla = True Then
                            msgError = "Error Espacio en blanco"
                            B_Error = True
                        End If
                        'Control de Caja
                        If Not IsNumeric(hojaEx.Cells(i, 2)) And hojaEx.Cells(i, 2) = "" Then
                            msgError = "Error en NRO_Caja"
                            B_Error = True
                         Else
                            ERRORCAJA = ""
                           If chkNoControlasEstado = 0 Then
'                            If ControlCaja(hojaEx.Cells(i, 2), fkcliente, ErrorCaja) = False Then
'                                msgError = ErrorCaja
'                                B_Error = True
'                            End If
                            End If
                        End If
                        'Control de documento
                        If Not IsNumeric(hojaEx.Cells(i, 3)) And hojaEx.Cells(i, 3) = "" And Not IsNumeric(hojaEx.Cells(i, 3)) Then
                           msgError = "Sector invalido"
                           B_Error = True
                        Else
                            If IsNumeric(.Cells(i, 3)) Then
                                ErrorDoc = ""
                                If TraerIndice(.Cells(i, 3), fkcliente, ErrorDoc) = "" Then
                                    msgError = ErrorDoc
                                    B_Error = True
                                End If
                            End If
                        End If

                            'Control de Fecha_desde
                            If .Cells(i, 5) <> "" Then
                               If Not IsDate(.Cells(i, 5)) Then
                                    msgError = "FECHA_DESDE"
                                    B_Error = True
                               Else
                                   If Len(.Cells(i, 5)) = 8 Or Len(.Cells(i, 5)) = 10 Then
                                        If CDate(.Cells(i, 5)) < "01/01/1940" Then
                                            msgError = "FECHA_DESDE"
                                            B_Error = True
                                         Else
                                            ValidarIndiceSecundario = True
                                        End If
                                    Else
                                        msgError = "FECHA_DESDE"
                                        B_Error = True
                                    End If
                               End If
                            End If

                            'Control Fecha Hasta
                            If .Cells(i, 6) <> "" Then
                               If Not IsDate(.Cells(i, 6)) Then
                                    msgError = "FECHA_HASTA"
                                    B_Error = True
                               End If
                            Else
                                If .Cells(i, 5) <> "" Then
                                    msgError = "Falta fecha hasta "
                                    B_Error = True
                                End If
                            End If
                            'Control  Desde

                            If .Cells(i, 7) <> "" Then
                               ValidarIndiceSecundario = True
                            End If


                            If ValidarIndiceSecundario = False Then
                               msgError = "Error índice Secundario"
                                B_Error = True
                            End If



                            If msgError <> "" Then
                                .Cells(i, 1) = msgError
                            Else
                                 Rem .Cells(i, 1) = ""
                            End If


                 End If


                msgError = ""
                lblTitulo.Caption = "Control registros"
                lblCantidadRegistros.Caption = i
                lblCantidadRegistros.Refresh
            Next
           Control_Excel_Cliente = B_Error
        End With
        hojaEx.Protect 21877471
        libroEx.Save
        libroEx.Close
        ApExcel.Quit
        Set hojaEx = Nothing
        Set libroEx = Nothing
        Set ApExcel = Nothing

        If ModificarNombre Then
            If B_Error = False Then
                FileSystem.FileCopy CommonDialog1.FileName, Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - 4) & " Control_OK " & " .xls"
                FileSystem.Kill CommonDialog1.FileName
            End If
        End If
        Exit Function
salir:
           hojaEx.Protect 21877471
           MsgBox "Error en la planilla"
           Control_Excel_Cliente = True
           libroEx.Close
            ApExcel.Quit
            Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
End Function

Private Function Percistencia_Excel_Cliente(PasoOrigen As String) As Boolean
        Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim NRO_CAJA, ID_referencia As Long
        Dim Cliente, i As Integer
        Dim NRO_DESDE, NRO_HASTA, Sql, Indice, msgError, Descripcion, FECHA_DESDE, FECHA_HASTA, LETRA_DESDE, LETRA_HASTA As String
        Dim strCodigo, Paso, s_NombreArchivo As String

       'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(PasoOrigen)
        Set hojaEx = libroEx.Worksheets.Item(1)

        s_NombreArchivo = "'" & NombreArchivo(PasoOrigen) & "'"



         Dim C_Error, C_Doc, C_Caja, C_Descripcion, C_Fecha_desde, C_Fecha_hasta, C_Desde, C_Hasta As Integer

    '
'

        With hojaEx
            For i = 7 To 6000
               'Control de fin de rows
                If Trim(Cells(i, 2)) = "" And Trim(Cells(i, 3)) = "" Then
                    Exit For
                End If

                Cliente = 4


                'Caja
                 NRO_CAJA = hojaEx.Cells(i, 2)

                'Documento
                 Indice = "'" & TraerIndice(.Cells(i, 3), CInt(Cliente)) & "'"

                'Descripcion
                If .Cells(i, 4) <> "" Then
                   Descripcion = "'" & Trim(UCase(Replace(.Cells(i, 4), "'", "´"))) & "'"
                Else
                   Descripcion = "Null"
                End If

                'Control de Fecha_desde
                If .Cells(i, 5) <> "" Then
                    FECHA_DESDE = Format(.Cells(i, 5), "dd/mm/yyyy")
                Else
                    FECHA_DESDE = "Null"
                End If

                'Control Fecha Hasta
                If .Cells(i, 6) <> "" Then
                   FECHA_HASTA = Format(.Cells(i, 6), "DD/MM/YYYY")
                Else
                    FECHA_HASTA = "Null"
                End If

                'Desde
                If .Cells(i, 7) <> "" Then
                    If IsNumeric(.Cells(i, 7)) Then
                        NRO_DESDE = .Cells(i, 7)
                        LETRA_DESDE = "Null"
                    Else
                        NRO_DESDE = "Null"
                        LETRA_DESDE = "'" & UCase(Trim(.Cells(i, 7))) & "'"
                    End If
                Else
                    NRO_DESDE = "Null"
                    LETRA_DESDE = "Null"
                End If

                'Hasta
                If .Cells(i, 8) <> "" Then
                    If IsNumeric(.Cells(i, 8)) Then
                        NRO_HASTA = .Cells(i, 8)
                        LETRA_HASTA = "Null"
                    Else
                        NRO_HASTA = "Null"
                        LETRA_HASTA = "'" & UCase(Trim(.Cells(i, 8))) & "'"
                    End If
                Else
                    NRO_HASTA = "Null"
                    LETRA_HASTA = "Null"
                End If

                ID_referencia = InsertarReferencias(CInt(Cliente), CLng(NRO_CAJA), CStr(Indice), CStr(Descripcion) _
                , CStr(FECHA_DESDE), CStr(FECHA_HASTA), CStr(NRO_DESDE), CStr(NRO_HASTA), CStr(LETRA_DESDE), CStr(LETRA_HASTA) _
                , "Null", CStr(s_NombreArchivo), 0, "Null", 99)


                     Rem Debug.Assert I < 200
                lblTitulo.Caption = "Control Grabación:"
                lblCantidadRegistros.Caption = i
                lblCantidadRegistros.Refresh

             Next
        End With

        MsgBox "La grabación de realizo con exito", vbInformation

        libroEx.Save
        libroEx.Close
        ApExcel.Quit
        Set hojaEx = Nothing
        Set libroEx = Nothing
        Set ApExcel = Nothing
        FileSystem.FileCopy CommonDialog1.FileName, Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - 4) & " Procesado" & " .xls"
        FileSystem.Kill CommonDialog1.FileName
        
End Function

Private Function Percistencia_Excel_Carga(PasoOrigen As String) As Boolean
'        Dim ApExcel As Excel.Application
'        Dim libroEx As Excel.Workbook
'        Dim hojaEx As Excel.Worksheet
'        Dim NRO_CAJA, ID_referencia, ID_SQL   As Long
'        Dim Cliente, i As Integer
'        Dim NRO_DESDE, NRO_HASTA, ControlExcel, sql, Indice, msgError, Descripcion, FECHA_DESDE, FECHA_HASTA, LETRA_DESDE, LETRA_HASTA As String
'        Dim strCodigo, paso, s_NombreArchivo, nombreArchivoOrigen As String
'        Dim NRO_CAJA_ANTERIOR As Long
'        Dim ContadorFilasBlanco As Integer
'        'abrir hoja excel
'        Set ApExcel = New Excel.Application
'        Set libroEx = Excel.Workbooks.Open(PasoOrigen)
'        Set hojaEx = libroEx.Worksheets.Item(1)
'       nombreArchivoOrigen = ""
'        s_NombreArchivo = "'" & NombreArchivo(PasoOrigen) & "'"
'
'        Dim C As Integer
'
'        Dim Control_Columnas  As Boolean
'
'            Dim C_Error, C_Doc, C_Caja, C_Descripcion, C_Fecha_desde, C_Fecha_hasta, C_Desde, C_Hasta As Integer
'    ' Control de columnas
'
'    With hojaEx
'    For C = 1 To 10
'        Control_Columnas = False
'        If UCase(.Cells(1, C)) = "ERROR" Then
'            C_Error = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "PROCESO" Then
'            C_Doc = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "CAJA" Then
'            C_Caja = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "DESCRIPCION" Or UCase(.Cells(1, C)) = "DESCRIPCIÓN" Then
'            C_Descripcion = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "FECHA DESDE" Then
'            C_Fecha_desde = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "FECHA HASTA" Then
'            C_Fecha_hasta = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "DESDE" Then
'            C_Desde = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) = "HASTA" Then
'            C_Hasta = C
'            Control_Columnas = True
'        End If
'        If UCase(.Cells(1, C)) <> "" And Control_Columnas = False Then
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            MsgBox "Error en nombre de columna " & Cells(1, C)
'            Exit Function
'        End If
'    Next
'
'
'
'
'
'
'
'
'            For i = 2 To 5000
'               'Control de fin de rows
'            If hojaEx.Cells(i, 2) = "" And .Cells(i, 3) = "" And .Cells(i, 4) = "" Then
'                        ContadorFilasBlanco = ContadorFilasBlanco + 1
'                        If ContadorFilasBlanco > 15 Then
'                            Exit For
'                        End If
'                        GoTo Saltar
'                    End If
'                'Ver
'                If .Cells(i, 1) <> "" Then
'                    ControlExcel = "'" & .Cells(i, 1) & "'"
'                Else
'                    ControlExcel = "Null"
'                End If
'
'                'Cliente
'                Cliente = ctlClientes.Valor
'
'                'Caja
'                 NRO_CAJA = hojaEx.Cells(i, C_Caja)
'
'                'Documento
'                If IsNumeric(.Cells(i, C_Doc)) Then
'                    Indice = "'" & TraerIndice(.Cells(i, C_Doc), ctlClientes.Valor) & "'"
'                     Descripcion = ""
'                 Else
'                    Indice = "'001'"
'                    Descripcion = "Sector: " & Trim(UCase(Replace(.Cells(i, 4), "'", "´")))
'                 End If
'
'                'Descripcion
'                If .Cells(i, C_Descripcion) <> "" Then
'                    If Descripcion = "" Then
'                        Descripcion = "'" & Trim(UCase(Replace(.Cells(i, C_Descripcion), "'", "´"))) & "'"
'                    Else
'                        Descripcion = "'" & Descripcion & " " & Trim(UCase(Replace(.Cells(i, C_Descripcion), "'", "´"))) & "'"
'                    End If
'                Else
'                   If Descripcion = "" Then
'                        Descripcion = "Null"
'                    Else
'                        Descripcion = "'" & Descripcion & "'"
'                    End If
'                End If
'
'                'Control de Fecha_desde
'                If .Cells(i, C_Fecha_desde) <> "" Then
'                    FECHA_DESDE = "'" & Format(.Cells(i, C_Fecha_desde), "dd/mm/yyyy") & "'"
'                Else
'                    FECHA_DESDE = "Null"
'                End If
'
'                'Control Fecha Hasta
'                If .Cells(i, C_Fecha_hasta) <> "" Then
'                   FECHA_HASTA = "'" & Format(.Cells(i, C_Fecha_hasta), "dd/mm/yyyy") & "'"
'                Else
'                   FECHA_HASTA = "Null"
'                End If
'
'                'NRO_Desde
'                If .Cells(i, C_Desde) <> "" Then
'                    NRO_DESDE = .Cells(i, C_Desde)
'                Else
'                    NRO_DESDE = "NULL"
'                End If
'
'                'NRO_Hasta
'                If .Cells(i, C_Hasta) <> "" Then
'                    NRO_HASTA = .Cells(i, C_Hasta)
'                Else
'                    NRO_HASTA = "NULL"
'                End If
'
'                 'Letra_desde
'                 If .Cells(i, C_Desde) <> "" Then
'                    LETRA_DESDE = "'" & .Cells(i, C_Desde) & "'"
'                 Else
'                   LETRA_DESDE = "Null"
'                 End If
'
'
'                 'Letra_Hasta
'                 If .Cells(i, C_Hasta) <> "" Then
'                    LETRA_HASTA = "'" & .Cells(i, C_Hasta) & "'"
'                 Else
'                    LETRA_HASTA = "Null"
'                 End If
'
'
'                If Trim(.Cells(i, 26)) <> "" Then
'
'                     ID_SQL = CLng(Trim(.Cells(i, 26)))
'                     InsertarImagenes ID_SQL, CInt(.Cells(i, 2)), CInt(.Cells(i, 3)), 1, Format(Now, "DD/MM/YYYY")
'                      NRO_CAJA_ANTERIOR = NRO_CAJA
'                End If
'
'                ID_referencia = InsertarReferencias(ctlClientes.Valor, CLng(NRO_CAJA), CStr(Indice), CStr(Descripcion) _
'                , CStr(FECHA_DESDE), CStr(FECHA_HASTA), CStr(NRO_DESDE), CStr(NRO_HASTA), CStr(LETRA_DESDE), CStr(LETRA_HASTA), "Null" _
'                 , CStr(s_NombreArchivo), CLng(ID_SQL), CStr(ControlExcel), ctlUsuario.Valor)
'                .Cells(i, 27) = ID_referencia
'
'                lblTitulo.Caption = "Control Grabación:"
'                lblTitulo.Refresh
'                lblCantidadRegistros.Caption = i
'                lblCantidadRegistros.Refresh
'Saltar:
'             Next
'        End With
'
'        MsgBox "La grabación de realizo con exito " & s_NombreArchivo, vbInformation
'
'        libroEx.Save
'        libroEx.Close
'        ApExcel.Quit
'        Set hojaEx = Nothing
'        Set libroEx = Nothing
'        Set ApExcel = Nothing
'        FileSystem.FileCopy PasoOrigen, Mid(PasoOrigen, 1, Len(PasoOrigen) - 4) & " Procesado" & " .xls"
'        FileSystem.Kill PasoOrigen
'
End Function

Private Function Percistencia_Excel_Legajos(PasoOrigen As String, ModificarNombre As Boolean) As Boolean
     Dim ApExcel As Excel.Application
     Dim libroEx As Excel.Workbook
     Dim hojaEx As Excel.Worksheet
     Dim NRO_CAJA, ID_referencia, ID_SQL   As Long
     Dim B_Error  As Boolean
     Dim Cliente, i As Integer
     Dim NRO_DESDE, NRO_HASTA, ControlExcel, Sql, Indice, msgError, Descripcion, FECHA_DESDE, FECHA_HASTA, LETRA_DESDE, LETRA_HASTA As String
     Dim strCodigo, Paso, s_NombreArchivo, nombreArchivoOrigen As String
     
     'abrir hoja excel
     Set ApExcel = New Excel.Application
     Set libroEx = Excel.Workbooks.Open(PasoOrigen)
     Set hojaEx = libroEx.Worksheets.Item(1)
    
     s_NombreArchivo = "'" & NombreArchivo(PasoOrigen) & "'"
        
        
        
        
    Dim C As Integer
 
    
    

    If hojaEx.Name <> "Carga_Legajos" Then
        MsgBox "El nombre de la planilla no es el correcto", vbInformation
        libroEx.Close
        ApExcel.Quit
        Set hojaEx = Nothing
        Set libroEx = Nothing
        Set ApExcel = Nothing
        Exit Function
    End If
    B_Error = False

With hojaEx
    Cells(1, 19) = "ID_Imagen"
    Cells(1, 20) = "ID_Oracle"
    For i = 2 To 10000
       'Control de fin de rows
       If Cells(i, 2) = "" And Cells(i, 3) = "" Then
            Exit For
       End If
          
        'Indice
         Indice = "'" & TraerIndice(.Cells(i, 4), .Cells(i, 2)) & "'"
        
        
        If Not (nombreArchivoOrigen = Trim(.Cells(i, 11))) Then
            nombreArchivoOrigen = Trim(.Cells(i, 11))
            ID_SQL = ValidarImagenSql("\" & Trim(.Cells(i, 10)) & "\" & Trim(.Cells(i, 11)) & ".tif")
            InsertarImagenes ID_SQL, CInt(.Cells(i, 2)), CInt(.Cells(i, 3)), 1, Format(Now, "DD/MM/YYYY")
        End If
        
        For C = 5 To 8
            If hojaEx.Cells(i, C) <> "" Then
                ID_referencia = InsertarReferencias(CInt(.Cells(i, 2)), CLng(.Cells(i, 3)), CStr(Indice), "Null", _
                  "Null", "Null", "Null", "Null", "Null", "Null", Trim(hojaEx.Cells(i, C)), _
                 CStr(s_NombreArchivo), CLng(ID_SQL), CStr(s_NombreArchivo), "'Sistema_Legajos'")
                .Cells(i, 20) = ID_referencia
             End If
        Next
        lblTitulo.Caption = "Control Grabación:"
                lblTitulo.Refresh
                lblCantidadRegistros.Caption = i
                lblCantidadRegistros.Refresh
    Next

End With

libroEx.Save
libroEx.Close
ApExcel.Quit
Set hojaEx = Nothing
Set libroEx = Nothing
Set ApExcel = Nothing
    If ModificarNombre Then
        If B_Error = False Then
            FileSystem.FileCopy PasoOrigen, Mid(PasoOrigen, 1, Len(PasoOrigen) - 4) & " Procesaedo " & " .xls"
            FileSystem.Kill PasoOrigen
        End If
    End If
    
End Function
Private Function Percistencia_Excel_Documento(PasoOrigen As String, ModificarNombre As Boolean) As Boolean
        
'
'        'Definiciones Excel
'        Dim ApExcel As Excel.Application
'        Dim libroEx As Excel.Workbook
'        Dim hojaEx As Excel.Worksheet
'        Set ApExcel = New Excel.Application
'        Set libroEx = Excel.Workbooks.Open(PasoOrigen)
'        Set hojaEx = libroEx.Worksheets.Item(2)
'
'        Dim B_Error As Boolean
'        Dim i As Integer
'        Dim msgError  As String
'        'abrir hoja excel
'        'Variables del docuemneto
'        Dim Documento_Actual As String
'        Dim Indice_Actual As String
'        Dim sql  As String
'        Dim ID_Referencias As String
'        Dim Error_Documento As String
'        Dim COD_CLIENTE As Integer
'        Dim s_NombreArchivo As String
'        Dim FechaSistema As String
'        s_NombreArchivo = "'" & NombreArchivo(PasoOrigen) & "'"
'       Percistencia_Excel_Documento = False
'       FechaSistema = SysDateMinutoSegundo
'        'Variables Tipo Pamilla
'        Dim C_ID As Integer
'        Dim C_Documento As Integer
'        Dim C_Error As Integer
'        Dim C_Caja As Integer
'        Dim C_Ok As Integer
'        Dim Row_Inicial As Integer
'        Dim Contar_Vacias  As Integer
'        Dim Indice_Actual_Control As String
'        Dim UpdateReferencias As Boolean
'        UpdateReferencias = True
'
'
'
'        On Error GoTo salir:
'               ConBasa.BeginTrans
'
'        'Control de Nombre de planilla
'        Select Case hojaEx.Name
'        Case "Cajas"
'            If Trim(Cells(2, 3)) <> "REFERENCIAS POR CAJAS" Then
'                MsgBox "Error en el formato", vbInformation
'                GoTo salir
'            Else
'                C_ID = 4
'                C_Documento = 2
'                C_Error = 5
'                C_Caja = 1
'                C_Ok = 6
'                Row_Inicial = 4
'            End If
'        Case "Referencias"
'            If Trim(Cells(1, 3)) <> "Referencia por indice" Then
'                MsgBox "Error en el formato", vbInformation
'                GoTo salir
'            Else
'                C_ID = 12
'                C_Documento = 1
'                C_Error = 13
'                C_Caja = 2
'                C_Ok = 13
'                Row_Inicial = 5
'            End If
'
'
'        Case "Inconsistencias"
'            If Trim(Cells(1, 3)) <> "Planilla de Inconsistencias" Then
'                MsgBox "Error en el formato", vbInformation
'                GoTo salir
'            Else
'                C_ID = 12
'                C_Documento = 1
'                C_Error = 13
'                C_Caja = 2
'                C_Ok = 13
'                Row_Inicial = 4
'            End If
'
'
'        Case Else
'            MsgBox "El nombre de la planilla no es el correcto", vbInformation
'            GoTo salir
'        End Select
'
'        'Control de Cliente
'        If IsNull(ctlClientes.Valor) Then
'            MsgBox "Ingrese el cliente", vbCritical
'            GoTo salir
'        Else
'            COD_CLIENTE = ctlClientes.Valor
'        End If
'
'    'Control de datos
'
'    With hojaEx
'            For i = Row_Inicial To 1500
'               Error_Documento = ""
'               ID_Referencias = 0
'
'               If Not (Len(Cells(i, C_Caja)) = 0 And Len(Cells(i, C_Documento)) = 0) Then
'
'                            Select Case hojaEx.Name
'                            Case "Cajas"
'                                        If Not IsNumeric(Cells(i, C_Caja)) Or Cells(i, C_Caja) = "" Then
'                                            If Cells(i, C_Documento) > 0 And IsNumeric(Cells(i, C_Documento)) Then
'                                                Contar_Vacias = 0
'                                                Indice_Actual_Control = TraerIndice(CInt(Cells(i, C_Documento)), COD_CLIENTE)
'                                                If Indice_Actual_Control = "" Then
'                                                    Error_Documento = Indice_Actual_Control
'                                                End If
'                                                ID_Referencias = Control_ID_Referencias(COD_CLIENTE, CLng(Replace(Cells(i, C_ID), "ID:", "")))
'                                                If ID_Referencias = "" Then
'                                                    Error_Documento = "El ID Referencia No corresponde"
'                                                End If
'                                            Else
'                                                Error_Documento = "Error En Asignacion de doc"
'                                            End If
'                                            UpdateReferencias = True
'                                        Else
'                                            UpdateReferencias = False
'                                        End If
'
'
'
'                            Case "Referencias"
'                                UpdateReferencias = True
'                                If IsNumeric(Cells(i, C_Documento)) = True Then
'                                         Contar_Vacias = 0
'                                         UpdateReferencias = True
'                                         Indice_Actual_Control = TraerIndice(CInt(Cells(i, C_Documento)), COD_CLIENTE)
'                                         If Indice_Actual_Control = "" Then
'                                             Error_Documento = Indice_Actual_Control
'                                         End If
'                                         ID_Referencias = Control_ID_Referencias(COD_CLIENTE, CLng(Cells(i, C_ID)))
'                                         If ID_Referencias = "" Then
'                                             Error_Documento = "El ID Referencia No corresponde"
'                                         End If
'                                 Else
'                                     If IsNumeric(Cells(i, C_Caja)) = False Then
'                                          Error_Documento = "Error En Asignacion de doc"
'                                     End If
'                                     UpdateReferencias = False
'                                 End If
'
'                            Case "Inconsistencias"
'                                UpdateReferencias = True
'                                If IsNumeric(Cells(i, C_Documento)) = True Then
'                                         Contar_Vacias = 0
'                                         UpdateReferencias = True
'                                         Indice_Actual_Control = TraerIndice(CInt(Cells(i, C_Documento)), COD_CLIENTE)
'                                         If Indice_Actual_Control = "" Then
'                                             Error_Documento = Indice_Actual_Control
'                                         End If
'                                         ID_Referencias = Control_ID_Referencias(COD_CLIENTE, CLng(Cells(i, C_ID)))
'                                         If ID_Referencias = "" Then
'                                             Error_Documento = "El ID Referencia No corresponde"
'                                         End If
'                                 Else
'                                     If IsNumeric(Cells(i, C_Caja)) = False Then
'                                          Error_Documento = "Error En Asignacion de doc"
'                                     End If
'                                     UpdateReferencias = False
'                                 End If
'
'
'
'                            End Select
'
'
'
'                                 If Error_Documento = "" And UpdateReferencias = True Then
'                                     sql = "   Update REFERENCIAS"
'                                     sql = sql & vbCrLf & " SET INDICE = '" & Indice_Actual_Control & "',"
'                                     sql = sql & vbCrLf & " PASOARCHIVO = " & s_NombreArchivo & ", "
'                                     sql = sql & vbCrLf & " FECHA_MODIFICACION =" & FechaSistema & ","
'                                     sql = sql & vbCrLf & " USUARIO_MODIFICACION = 'S_Cambio_indice'"
'                                     sql = sql & vbCrLf & " Where COD_ID_REFERENCIA = " & ID_Referencias
'                                     sql = sql & vbCrLf & " And COD_CLIENTE = " & COD_CLIENTE
'                                     ExecutarSql sql
'                                 End If
'                     'Control de Error
'                     If Error_Documento <> "" Then
'                         Cells(i, C_Error) = Error_Documento
'                         B_Error = True
'                     Else
'                         If ID_Referencias <> 0 Then
'                             Cells(i, C_Ok) = "OK"
'                         Else
'                             Cells(i, C_Ok) = "Caja"
'                         End If
'                     End If
'                lblTitulo.Caption = "Control Grabación:"
'                lblCantidadRegistros.Caption = i
'                lblCantidadRegistros.Refresh
'
'                Else
'                    Contar_Vacias = Contar_Vacias + 1
'                    If Contar_Vacias > 5 Then
'                        Exit For
'                    End If
'                End If
'
'            Next
'        End With
'
'            libroEx.Save
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            'Control Base
'            If B_Error = False Then
'                ConBasa.CommitTrans
'                MsgBox "La modificacion se realizo con exito", vbInformation
'                Percistencia_Excel_Documento = True
'              Else
'                ConBasa.RollbackTrans
'                MsgBox "Verifique la planilla e inténtelo nuevamente", vbCritical
'            End If
'            'NombreArchivo
'            If ModificarNombre Then
'                If B_Error = False Then
'                    FileSystem.FileCopy PasoOrigen, Mid(PasoOrigen, 1, Len(PasoOrigen) - 4) & " Procesado el " & Format(Now, "dd_mm_yyyy") & ".xls"
'                    FileSystem.Kill PasoOrigen
'                End If
'            End If
'        Exit Function
'salir:
'            Percistencia_Excel_Documento = False
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            ConBasa.RollbackTrans
'            MsgBox "Error en la planilla" & vbCrLf & Err.Description, vbCritical
End Function

Private Sub cmdControlCarga_Click()
'        Dim Paso As String
'        CommonDialog1.ShowOpen
'        Paso = CommonDialog1.FileName
'        If Control_Excel_Carga(Paso, True) Then
'            MsgBox "Existen errores en la planilla", vbCritical, "Control de Excel"
'        Else
'            MsgBox "NO existen errores en la planilla", vbInformation, "Control de Excel"
'        End If
End Sub


Private Sub cmdControlIndices_Click()
        Dim Paso As String
        CommonDialog1.FileName = "\\Serverbackup_1\E\Usuarios\Clientes\Servicios\1- Ecogas\Incosistencias\*.XLS"
        CommonDialog1.ShowOpen
    If CommonDialog1.CancelError <> True Then
        Paso = CommonDialog1.FileName
        If Control_Excel_Cliente_Documento(Paso, True) Then
            MsgBox "Existe un errores", vbCritical, "Control Carga"
        Else
            MsgBox "NO Existe errores", vbCritical, "Control Carga"
            cmdPercistenciaDocumetos.Enabled = True
        End If
    End If
End Sub

Private Sub cmdBorrarTablaRecetas_Click()

End Sub

Private Sub cmdControlLegajos_Click()
On Error GoTo salir
 Dim Paso As String

    
     

    Paso = InputBox("ingrese el paso", Paso, "Z:\Administracion\Referencias\")


       
        CommonDialog1.FileName = Paso & "*.xls"

        CommonDialog1.ShowOpen
        Paso = CommonDialog1.FileName

        Control_Referencias (Paso)


salir:
    
End Sub

Private Sub cmdLegajosOsep_Click()
Dim Paso As String
        CommonDialog1.FileName = "\\Serverbackup_1\E\Usuarios\Clientes\Público\1- Osep\Envio de expedientes"
        CommonDialog1.ShowOpen
        Paso = CommonDialog1.FileName
        If Paso = "\\Serverbackup_1\E\Usuarios\Clientes\Público\1- Osep\Envio de expedientes" Then
            Exit Sub
        End If
        PercistenciaLegajosOsde Paso
End Sub

Private Sub cmdPercistenciaCarga_Click()
'    Dim Paso As String
'    On Error GoTo Salir
'        CommonDialog1.FileName = "c:\*.XLS"
'        CommonDialog1.ShowOpen
'        Paso = CommonDialog1.FileName
'    If Control_Excel_Carga(Paso, False) Then
'        MsgBox "Existe un errores" & vbCr & CommonDialog1.FileName, vbCritical, "Control Carga"
'    Else
'        Percistencia_Excel_Carga Paso
'    End If
'Salir:
End Sub

Private Sub cmdPercistenciaCliente_Click()
    Dim Paso As String
On Error GoTo salir
        CommonDialog1.FileName = "c:\*.xls"
        CommonDialog1.ShowOpen
        Paso = CommonDialog1.FileName
        If Dir(CommonDialog1.FileName) = "" Then
        Exit Sub
        End If
        If Control_Excel_Cliente(Paso, False) Then
            MsgBox "Existe un errores"
        Else
            Percistencia_Excel_Cliente Paso
        End If
        Exit Sub
salir:
    MsgBox "Error en proceso"
End Sub


Private Sub cmdOsepRecetas_Click()

        Dim MyPath As String
        Dim MyName As String
        Dim Paso As String
        
        ExecutarSql "DELETE FROM TEM_OSEP_RECETAS"
        
        Paso = "Z:\Administracion\Referencias\FARMALINK\" & txtQuincenaMesAño.Text & "\"
        MyPath = "Z:\Administracion\Referencias\FARMALINK\" & txtQuincenaMesAño.Text & "\" & "*.xls"   ' Set the path.
        MyName = Dir(MyPath)   ' Retrieve the first entry.
        Do While MyName <> ""   ' Start the loop.
           LeerExcelRecetas Paso & MyName, Mid(txtQuincenaMesAño.Text, 3), Mid(txtQuincenaMesAño.Text, 1, 2), 123, MyName
           MyName = Dir()   ' Get next entry.
        Loop
        CrearPlanilla
End Sub

Private Sub cmdPercistenciaDocumetos_Click()
'   Dim paso As String
'On Error GoTo salir
'        CommonDialog1.FileName = "\\Serverbackup_1\E\Usuarios\Clientes\Servicios\1- Ecogas\*.XLS"
'        If IsNull(ctlClientes.Valor) Then
'            MsgBox "Ingrese el cliente", vbCritical
'            Exit Sub
'         End If
'
'
'        CommonDialog1.ShowOpen
'    If CommonDialog1.CancelError = True Then
'        paso = CommonDialog1.FileName
'
'        'Control de Cliente
'
'
'        If Percistencia_Excel_Documento(paso, True) = False Then
'            MsgBox "Existe errores en la planilla", vbCritical, "Control Indice"
'         End If
'    End If
'    Exit Sub
'salir:
'    If Err.Number = 32755 Then
'        MsgBox "Operación Cancelada", vbInformation
'    Else
'        MsgBox Err.Description
'    End If
    
End Sub

Private Sub cmdPlanillaManual_Click()
cmdPlanillaManual2

'            Dim comPlanilla As New ADODB.Connection
'            Dim RsPlanilla As New ADODB.Recordset
'            Dim Sql As String
'            Dim i As Integer
'            Dim FECHA_DESDE(4) As String
'            Dim FECHA_HASTA(4) As String
'            Dim NUMERO_DESDE(4) As String
'            Dim NUMERO_HASTA(4) As String
'            Dim INDICE(4) As String
'            Dim DESCRIPCION(4) As String
'            comPlanilla.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=Z:\Sistemas\Basa\BaseTeleform\Referencias.mdb"
'
'
'
'
'    Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'
'
'
'
'
'
'    Dim C_Error As Integer
'    Dim C_Caja As Integer
'    Dim C_Indice As Integer
'    Dim C_Etiqueta As Integer
'    Dim C_Fecha_desde As Integer
'    Dim C_Fecha_hasta As Integer
'    Dim C_N°_Desde As Integer
'    Dim C_N°_Hasta As Integer
'    Dim C_Letra_Desde As Integer
'    Dim C_Letra_Hasta As Integer
'    Dim C_Descripcion As Integer
'    Dim C_CSID As Integer
'    Dim C_TIPO As Integer
'    Dim R As String
'    Dim ErrorGeneral As Boolean
'    Dim strError As String
'    Dim FechaHora As String
'    Dim NombreArchivo As String
'    Dim directorio As String
'    Dim ControlIndice As Boolean
'    FechaHora = Trim(Format(Now, "hhmmss"))
'
'
'    C_Error = 1
'    C_Caja = 2
'    C_Indice = 4
'    C_Etiqueta = 3
'    C_Fecha_desde = 6
'    C_Fecha_hasta = 7
'    C_N°_Desde = 8
'    C_N°_Hasta = 9
'    C_Letra_Desde = 10
'    C_Letra_Hasta = 11
'    C_Descripcion = 5
'    C_TIPO = 16
'    C_CSID = 17
'
'
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Open("\\222.15.19.251\basa\Administracion\Referencias\" & "Planilla Modelo.xls", , True)
'    Set hojaEx = libroEx.Worksheets.Item(1)
'
'
'Sql = " SELECT Suspense_File, BatchNo, BatchPgNo, CAJA_Nº,ENVIO_CAJAS, USUARIO , CSID "
'Sql = Sql & vbCrLf & ", INDICE_1,DESCRIPCION_1,DIA_DESDE_1, MES_DESDE_1 , AÑO_DESDE_1 ,DIA_HASTA_1, MES_HASTA_1, AÑO_HASTA_1, NUMERO_DESDE_1 ,NUMERO_HASTA_1"
'Sql = Sql & vbCrLf & " ,INDICE_2,IDEM_INDICE_2, DESCRIPCION_2,IDEM_DETALLE_2,DIA_DESDE_2, MES_DESDE_2 , AÑO_DESDE_2 ,DIA_HASTA_2, MES_HASTA_2, AÑO_HASTA_2, NUMERO_DESDE_2 ,NUMERO_HASTA_2"
'Sql = Sql & vbCrLf & " ,INDICE_3,IDEM_INDICE_3, DESCRIPCION_3,IDEM_DETALLE_3,DIA_DESDE_3, MES_DESDE_3 , AÑO_DESDE_3 ,DIA_HASTA_3, MES_HASTA_3, AÑO_HASTA_3, NUMERO_DESDE_3 ,NUMERO_HASTA_3"
'Sql = Sql & vbCrLf & " ,INDICE_4,IDEM_INDICE_4, DESCRIPCION_4,IDEM_DETALLE_4,DIA_DESDE_4, MES_DESDE_4 , AÑO_DESDE_4 ,DIA_HASTA_4, MES_HASTA_4, AÑO_HASTA_4, NUMERO_DESDE_4 ,NUMERO_HASTA_4,ENVIO_CAJAS"
'Sql = Sql & vbCrLf & " From REFERENCIAS"
'Sql = Sql & vbCrLf & " GROUP BY Suspense_File, BatchNo, BatchPgNo, CAJA_Nº,ENVIO_CAJAS, USUARIO , CSID "
'Sql = Sql & vbCrLf & " ,INDICE_1,DESCRIPCION_1,DIA_DESDE_1, MES_DESDE_1 , AÑO_DESDE_1 ,DIA_HASTA_1, MES_HASTA_1, AÑO_HASTA_1, NUMERO_DESDE_1 ,NUMERO_HASTA_1"
'Sql = Sql & vbCrLf & " ,INDICE_2,IDEM_INDICE_2, DESCRIPCION_2,IDEM_DETALLE_2,DIA_DESDE_2, MES_DESDE_2 , AÑO_DESDE_2 ,DIA_HASTA_2, MES_HASTA_2, AÑO_HASTA_2, NUMERO_DESDE_2 ,NUMERO_HASTA_2"
'Sql = Sql & vbCrLf & " ,INDICE_3,IDEM_INDICE_3, DESCRIPCION_3,IDEM_DETALLE_3,DIA_DESDE_3, MES_DESDE_3 , AÑO_DESDE_3 ,DIA_HASTA_3, MES_HASTA_3, AÑO_HASTA_3, NUMERO_DESDE_3 ,NUMERO_HASTA_3"
'Sql = Sql & vbCrLf & " ,INDICE_4,IDEM_INDICE_4, DESCRIPCION_4,IDEM_DETALLE_4,DIA_DESDE_4, MES_DESDE_4 , AÑO_DESDE_4 ,DIA_HASTA_4, MES_HASTA_4, AÑO_HASTA_4, NUMERO_DESDE_4 ,NUMERO_HASTA_4, ENVIO_CAJAS"
'Sql = Sql & vbCrLf & " Having BatchNo = " & InputBox("Ingrese el numero de LOTE TELEFORM")
'Sql = Sql & vbCrLf & " ORDER BY BatchNo, BatchPgNo;"
'
'
'
'            RsPlanilla.Open Sql, comPlanilla
'            R = 7
'            If RsPlanilla.EOF Then
'                MsgBox "NO EXISTE EL LOTE "
'
'                libroEx.Close
'                ApExcel.Quit
'                Set hojaEx = Nothing
'                Set libroEx = Nothing
'                Set ApExcel = Nothing
'                Exit Sub
'            End If
'
'
'        ControlIndice = False
'            Do While Not RsPlanilla.EOF
'
'
'           Debug.Assert RsPlanilla!CAJA_Nº <> 875770
'
'            If i > 4 Then
'                i = 1
'                Exit For
'            End If
'
'                For i = 1 To 4
'                    If Not IsNull(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))) Then
'                        Rem MsgBox RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))
'                        FECHA_DESDE(i) = Format(Format(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_DESDE_" & CStr(i)), "00"), "DD/MM/YYYY")
'                        If Not IsNull(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i))) Then
'
'
'                           If RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)) = 0 Then
'                                FECHA_HASTA(i) = FECHA_DESDE(i)
'                           Else
'
'                                If Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") <> "00" Then
'                                    FECHA_HASTA(i) = Format(Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_HASTA_" & CStr(i)), "00"), "DD/MM/YYYY")
'                                Else
'                                    FECHA_HASTA(i) = FECHA_DESDE(i)
'                                End If
'                            End If
'                        Else
'                            FECHA_HASTA(i) = FECHA_DESDE(i)
'                        End If
'                    Else
'                        FECHA_DESDE(i) = ""
'                    End If
'                    If Not IsNull(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))) Then
'                        Rem MsgBox RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))
'                        NUMERO_DESDE(i) = Trim(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i)))
'                        If Format(RsPlanilla.Fields.Item("NUMERO_HASTA_" & CStr(i)), "") <> "" Then
'                            NUMERO_HASTA(i) = Trim(RsPlanilla.Fields.Item("NUMERO_HASTA_" & i))
'                        Else
'                             NUMERO_HASTA(i) = NUMERO_DESDE(i)
'                        End If
'                    Else
'                        NUMERO_DESDE(i) = ""
'                        NUMERO_HASTA(i) = ""
'                    End If
'
'                    If Not IsNull(RsPlanilla.Fields.Item("INDICE_" & CStr(i))) Then
'                        INDICE(i) = RsPlanilla.Fields.Item("INDICE_" & CStr(i))
'                        Else
'                            If i <> 1 Then
'                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_INDICE_" & CStr(i))) Then
'                                    INDICE(i) = INDICE(CStr(i - 1))
'                                 Else
'                                    If Not IsNull(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))) Or Not IsNull(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))) Or Not IsNull(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))) Then
'                                     INDICE(i) = "0"
'                                    Else
'                                     INDICE(i) = ""
'                                    End If
'                                End If
'                            Else
'                                If INDICE(i) = "" Then
'                                INDICE(i) = "0"
'                                Else
'                                End If
'
'                            End If
'
'                    End If
'
'                    If Not IsNull(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))) Then
'                        DESCRIPCION(i) = LCase(Trim(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))))
'                    Else
'                            If i <> 1 Then
'                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_DETALLE_" & CStr(i))) Then
'                                If RsPlanilla.Fields.Item("IDEM_DETALLE_" & CStr(i)) <> 0 Then
'
'                                    DESCRIPCION(i) = LCase(DESCRIPCION(CStr(i - 1)))
'                                 Else
'                                 DESCRIPCION(i) = ""
'                                 End If
'                                Else
'                                DESCRIPCION(i) = ""
'                                End If
'                            Else
'                                DESCRIPCION(i) = ""
'                            End If
'                    End If
'
'
'
'
'
'                If IsNull(RsPlanilla.Fields.Item("CAJA_Nº").value) Then
'                        directorio = "Error"
'                        hojaEx.Cells(R, C_Caja) = "Error"
'                        hojaEx.Cells(R, C_Indice) = ""
'                        hojaEx.Cells(R, C_Descripcion) = ""
'                        hojaEx.Cells(R, C_Fecha_desde) = ""
'                        hojaEx.Cells(R, C_Fecha_hasta) = ""
'                        hojaEx.Cells(R, C_N°_Desde) = ""
'                        hojaEx.Cells(R, C_N°_Hasta) = ""
'                        GoTo ERRORCAJA:
'                Else
'                        NombreArchivo = RsPlanilla.Fields.Item("CAJA_Nº").value & "_" & FechaHora & ".tif"
'               End If
'               Select Case RsPlanilla!ENVIO_CAJAS
'                        Case 1
'                            hojaEx.Cells(R, C_TIPO) = "Sin Refrencia"
'                            DESCRIPCION(i) = DESCRIPCION(i) & " //" & hojaEx.Cells(R, C_TIPO)
'                        Case 2
'                            hojaEx.Cells(R, C_TIPO) = "Envio Por correo"
'                            DESCRIPCION(i) = DESCRIPCION(i) & " // " & hojaEx.Cells(R, C_TIPO)
'                        Case 3
'                            hojaEx.Cells(R, C_TIPO) = "Referencia en planta"
'                            DESCRIPCION(i) = DESCRIPCION(i) & " // " & hojaEx.Cells(R, C_TIPO)
'                        Case 4
'                            hojaEx.Cells(R, C_TIPO) = "Cargar Legajos"
'                            DESCRIPCION(i) = DESCRIPCION(i) & " // " & hojaEx.Cells(R, C_TIPO)
'                        Case 5
'                            hojaEx.Cells(R, C_TIPO) = "Digitalizar"
'                            DESCRIPCION(i) = DESCRIPCION(i) & " " & hojaEx.Cells(R, C_TIPO)
'                        End Select
'
'
'              If INDICE(i) = "" And Trim(RsPlanilla.Fields.Item("CAJA_Nº").value) <> "" And hojaEx.Cells(R, C_TIPO) <> "" Then
'                    ControlIndice = True
'                    INDICE(i) = "1"
'                    DESCRIPCION(1) = DESCRIPCION(1) & " " & hojaEx.Cells(R, C_TIPO)
'              Else
'                  ControlIndice = False
'              End If
'
'
'                    If Trim(INDICE(i)) <> "" Then
'                        hojaEx.Cells(R, C_Caja) = RsPlanilla.Fields.Item("CAJA_Nº").value
'
'                        hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & Trim(RsPlanilla.Fields.Item("CAJA_Nº").value) & "\" & NombreArchivo
'                        hojaEx.Cells(R, C_Indice) = INDICE(i)
'                        Rem Debug.Print "Incide: " & Indice(i)
'                        hojaEx.Cells(R, C_Descripcion) = Trim(Replace(Replace(DESCRIPCION(i), vbCrLf, " "), vbCr, " "))
'
'                        If Trim(CStr(FECHA_DESDE(i))) <> "" Then
'                            If IsDate(FECHA_DESDE(i)) Then
'
'
'                                hojaEx.Cells(R, C_Fecha_desde).value = " " & CStr(FECHA_DESDE(i))
'
'                            Else
'                                hojaEx.Cells(R, C_Fecha_desde) = ""
'                            End If
'
'                        Else
'                            hojaEx.Cells(R, C_Fecha_desde).value = Format(FECHA_DESDE(i), "DD/MM/YYYY")
'                        End If
'                       Debug.Print "Caja : " & RsPlanilla.Fields.Item("CAJA_Nº").value & "//Incide: " & INDICE(i) & "//Fecha : " & CStr(FECHA_DESDE(i)) & "//dese:" & hojaEx.Cells(R, C_Descripcion) & " //Numero :" & NUMERO_DESDE(i)
'                       Rem  Debug.Print "Fecha : " & CStr(FECHA_DESDE(i))
'
'                       If Trim(CStr(FECHA_DESDE(i))) <> "" Then
'                            If IsDate(Format(FECHA_HASTA(i), "DD/MM/YYYY")) Then
'                                hojaEx.Cells(R, C_Fecha_hasta).value = " " & Format(FECHA_HASTA(i), "DD/MM/YYYY")
'                            Else
'                                hojaEx.Cells(R, C_Fecha_hasta).value = ""
'                            End If
'
'                       Else
'                            hojaEx.Cells(R, C_Fecha_hasta) = ""
'                       End If
'                       If Trim(NUMERO_DESDE(i)) <> "" Then
'                        hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(i)
'                        hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(i)
'                       Else
'                            hojaEx.Cells(R, C_N°_Desde) = ""
'                            hojaEx.Cells(R, C_N°_Hasta) = ""
'                            NUMERO_DESDE(i) = ""
'                            NUMERO_HASTA(i) = ""
'                       End If
'
'              If ControlIndice = True Then
'                 i = 5
'                ControlIndice = False
'              End If
'
'ERRORCAJA:
'
'
'
'                          hojaEx.Cells(R, C_CSID) = RsPlanilla.Fields.Item("CSID").value
'                          lblCantidadRegistros.Caption = RsPlanilla.Fields.Item("CSID").value
'
'                        R = R + 1
'                    End If
'                Next
'
'
'                If Dir("\\222.15.19.222\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3)) <> "" Then
'
'                If Not IsNull(RsPlanilla.Fields.Item("CAJA_Nº")) Then
'                        If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & CStr(RsPlanilla.Fields.Item("CAJA_Nº")), vbDirectory) = "" Then
'                            MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & CStr(RsPlanilla.Fields.Item("CAJA_Nº").value)
'                            FileCopy "\\222.15.19.222\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
'                         Else
'                            FileCopy "\\222.15.19.222\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
'                        End If
'
'                        Else
'                            Rem MsgBox "No se encontro La imagen " & RsPlanilla.Fields.Item("Suspense_File").value
'                        End If
'                  Else
'                        If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\", vbDirectory) = "" Then
'                            MkDir ("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\")
'                            FileCopy "\\TELEFORM\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\" & RsPlanilla.Index & Format(Now, "ddmmyyyy mmss") & ".tif"
'                         Else
'                            FileCopy "\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "Z:\Referencias" & "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\" & RsPlanilla.Index & Format(Now, "ddmmyyyy mmss") & ".tif"
'                        End If
'                   End If
'
'
'Proximo:
'
'
'                 RsPlanilla.MoveNext
'        Loop
'
'        Dim Paso_planilla As String
'
'        Paso_planilla = "\\222.15.19.251\basa\Administracion\Referencias\" & InputBox("Ingrese el nombre de la planilla") & "    " & Format(Now, "ddmmyyy hhss") & ".xls"
'
'                libroEx.SaveAs Paso_planilla
'                libroEx.Close
'                ApExcel.Quit
'                Set hojaEx = Nothing
'                Set libroEx = Nothing
'                Set ApExcel = Nothing
'
'                MsgBox "Terminado"
'            If MsgBox("Usted quiere ver la planilla", vbYesNo) = vbYes Then
'            Shell "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE " & Chr(34) & Paso_planilla & Chr(34), vbNormalFocus
'            End If


End Sub


Private Sub cmdPlanillaManual2()

            Dim comPlanilla As New ADODB.Connection
            Dim RsPlanilla As New ADODB.Recordset
            Dim Sql As String
            Dim i As Integer
            Dim Caja As String
            Dim FECHA_DESDE(4) As String
            Dim FECHA_HASTA(4) As String
            Dim NUMERO_DESDE(4) As String
            Dim NUMERO_HASTA(4) As String
            Dim LETRA_DESDE(4) As String
            Dim LETRA_HASTA(4) As String
            Dim Indice(4) As String
            Dim Descripcion(4) As String
            Dim TIPO_ENVIO As String
            Dim PasoFinal As String
            PasoFinal = "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\"
            Dim PasoInicial As String
            PasoInicial = "\\PCTELEMEMO1"
            
            
            
            
            
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet
    
    
    
        
   
    
    Dim C_Error As Integer
    Dim C_Caja As Integer
    Dim C_Indice As Integer
    Dim C_Etiqueta As Integer
    Dim C_Fecha_desde As Integer
    Dim C_Fecha_hasta As Integer
    Dim C_N°_Desde As Integer
    Dim C_N°_Hasta As Integer
    Dim C_Letra_Desde As Integer
    Dim C_Letra_Hasta As Integer
    Dim C_Descripcion As Integer
    Dim C_CSID As Integer
    Dim C_TIPO As Integer
    Dim R As String
    Dim ErrorGeneral As Boolean
    Dim strError As String
    Dim FechaHora As String
    Dim NombreArchivo As String
    Dim directorio As String
    Dim ControlIndice As Boolean
    FechaHora = Trim(Format(Now, "hhmmss"))
    
 
    C_Error = 1
    C_Caja = 2
    C_Indice = 4
    C_Etiqueta = 3
    C_Fecha_desde = 6
    C_Fecha_hasta = 7
    C_N°_Desde = 8
    C_N°_Hasta = 9
    C_Letra_Desde = 10
    C_Letra_Hasta = 11
    C_Descripcion = 5
    C_TIPO = 16
    C_CSID = 17
    
    
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Rem Set libroEx = Excel.Workbooks.Open("\\222.15.19.251\basa\Administracion\Referencias\" & "Planilla Modelo.xls", , True)
      Set libroEx = Excel.Workbooks.Open("c:\Planilla Modelo.xls", , True)
    
    Set hojaEx = libroEx.Worksheets.Item(1)
            
'   Dim celFecha As Excel.CellFormat
'   celFecha.NumberFormat
Dim lote As String

lote = InputBox("Ingrese el numero de LOTE TELEFORM")
If Not IsNumeric(lote) And Trim(lote) = "" Then
    MsgBox "NO es un numero "
     Exit Sub
End If


            Sql = " SELECT  Suspense_File, Form_Id , BatchNo, BatchPgNo, CAJA_Nº,ENVIO_CAJAS, USUARIO , CSID "
            Sql = Sql & vbCrLf & ", INDICE_1,DESCRIPCION_1,DIA_DESDE_1, MES_DESDE_1 , AÑO_DESDE_1 ,DIA_HASTA_1, MES_HASTA_1, AÑO_HASTA_1, NUMERO_DESDE_1 ,NUMERO_HASTA_1"
            Sql = Sql & vbCrLf & " ,INDICE_2,IDEM_INDICE_2, DESCRIPCION_2,IDEM_DETALLE_2,DIA_DESDE_2, MES_DESDE_2 , AÑO_DESDE_2 ,DIA_HASTA_2, MES_HASTA_2, AÑO_HASTA_2, NUMERO_DESDE_2 ,NUMERO_HASTA_2"
            Sql = Sql & vbCrLf & " ,INDICE_3,IDEM_INDICE_3, DESCRIPCION_3,IDEM_DETALLE_3,DIA_DESDE_3, MES_DESDE_3 , AÑO_DESDE_3 ,DIA_HASTA_3, MES_HASTA_3, AÑO_HASTA_3, NUMERO_DESDE_3 ,NUMERO_HASTA_3"
            Sql = Sql & vbCrLf & " ,INDICE_4,IDEM_INDICE_4, DESCRIPCION_4,IDEM_DETALLE_4,DIA_DESDE_4, MES_DESDE_4 , AÑO_DESDE_4 ,DIA_HASTA_4, MES_HASTA_4, AÑO_HASTA_4, NUMERO_DESDE_4 ,NUMERO_HASTA_4,ENVIO_CAJAS "
            Sql = Sql & vbCrLf & " ,INDICE_SECUNDARIO_1 , INDICE_SECUNDARIO_2, INDICE_SECUNDARIO_3, INDICE_SECUNDARIO_4"
            Sql = Sql & vbCrLf & " From TELEFORM_REFERENCIAS "
            Sql = Sql & vbCrLf & " GROUP BY Suspense_File,Form_Id , BatchNo, BatchPgNo, CAJA_Nº,ENVIO_CAJAS, USUARIO , CSID "
            Sql = Sql & vbCrLf & " ,INDICE_1,DESCRIPCION_1,DIA_DESDE_1, MES_DESDE_1 , AÑO_DESDE_1 ,DIA_HASTA_1, MES_HASTA_1, AÑO_HASTA_1, NUMERO_DESDE_1 ,NUMERO_HASTA_1"
            Sql = Sql & vbCrLf & " ,INDICE_2,IDEM_INDICE_2, DESCRIPCION_2,IDEM_DETALLE_2,DIA_DESDE_2, MES_DESDE_2 , AÑO_DESDE_2 ,DIA_HASTA_2, MES_HASTA_2, AÑO_HASTA_2, NUMERO_DESDE_2 ,NUMERO_HASTA_2"
            Sql = Sql & vbCrLf & " ,INDICE_3,IDEM_INDICE_3, DESCRIPCION_3,IDEM_DETALLE_3,DIA_DESDE_3, MES_DESDE_3 , AÑO_DESDE_3 ,DIA_HASTA_3, MES_HASTA_3, AÑO_HASTA_3, NUMERO_DESDE_3 ,NUMERO_HASTA_3"
            Sql = Sql & vbCrLf & " ,INDICE_4,IDEM_INDICE_4, DESCRIPCION_4,IDEM_DETALLE_4,DIA_DESDE_4, MES_DESDE_4 , AÑO_DESDE_4 ,DIA_HASTA_4, MES_HASTA_4, AÑO_HASTA_4, NUMERO_DESDE_4 ,NUMERO_HASTA_4, ENVIO_CAJAS"
            Sql = Sql & vbCrLf & " ,INDICE_SECUNDARIO_1 , INDICE_SECUNDARIO_2, INDICE_SECUNDARIO_3, INDICE_SECUNDARIO_4"
            Sql = Sql & vbCrLf & " Having BatchNo = " & lote
            Sql = Sql & vbCrLf & " ORDER BY BatchNo, BatchPgNo;"

            
            
            RsPlanilla.Open Sql, strConBasa
            R = 7
            If RsPlanilla.EOF Then
                MsgBox "NO EXISTE EL LOTE "
                
                libroEx.Close
                ApExcel.Quit
                Set hojaEx = Nothing
                Set libroEx = Nothing
                Set ApExcel = Nothing
                Exit Sub
            End If
            
     
        ControlIndice = False
        With RsPlanilla
            Do While Not .EOF
             TIPO_ENVIO = ""
             Caja = ""
            lblCantidadRegistros.Caption = !CSID
            lblCantidadRegistros.Refresh
          If Not IsNull(!CAJA_Nº) Then
                If IsNumeric(!CAJA_Nº) Then
                    If !CAJA_Nº > 10 Then
                      Caja = !CAJA_Nº
                    Else
                      Caja = !CSID
                    End If
                Else
                   Caja = !CSID
                End If
          Else
                Caja = !CSID
          End If
         
          
            Select Case !ENVIO_CAJAS
            Case 1
                TIPO_ENVIO = "//Sin Refrencia// "
            Case 2
                TIPO_ENVIO = "//Envio Por correo// "
            Case 3
                TIPO_ENVIO = "//Referencia en planta// "
            Case 4
                TIPO_ENVIO = "//Cargar Legajos// "
            Case 5
                TIPO_ENVIO = "//Digitalizar// "
           Case Else
           TIPO_ENVIO = ""
           End Select
            NombreArchivo = Caja & "_" & R & ".TIF"
            
            
            
            
            If Dir(PasoFinal & Caja, vbDirectory) = "" Then
                MkDir (PasoFinal & Caja)
                FileCopy PasoInicial & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), PasoFinal & Caja & "\" & NombreArchivo
            Else
                FileCopy PasoInicial & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), PasoFinal & Caja & "\" & NombreArchivo
            End If
         
         
     '_______________________ INICIO 1________________
        If Not IsNull(!INDICE_1) Then
           If IsNumeric(!INDICE_1) Then
               Indice(1) = !INDICE_1
            Else
               Indice(1) = "0"
           End If
        Else
           Indice(1) = "0"
        End If
        Descripcion(1) = ""
        If Not IsNull(!DESCRIPCION_1) Then
            Descripcion(1) = Trim(LCase(!DESCRIPCION_1))
            If !Form_Id = 189 And IsNumeric(Descripcion(1)) Then
               Descripcion(1) = Indice_Secundario_Disco(CInt(Descripcion(1)))
            End If
        Else
            Descripcion(1) = ""
        End If
        
        FECHA_DESDE(1) = FECHA_DESDE_SOLUCION(!DIA_DESDE_1, !MES_DESDE_1, !AÑO_DESDE_1)
        FECHA_HASTA(1) = FECHA_HASTA_SOLUCION(!DIA_HASTA_1, !MES_HASTA_1, !AÑO_HASTA_1)
        

        
        
        
        If Not IsNull(!NUMERO_DESDE_1) Then
            If IsNumeric(!NUMERO_DESDE_1) Then
                NUMERO_DESDE(1) = !NUMERO_DESDE_1
                LETRA_DESDE(1) = ""
            Else
                NUMERO_DESDE(1) = ""
                LETRA_DESDE(1) = !NUMERO_DESDE_1
            End If
        Else
            NUMERO_DESDE(1) = ""
            LETRA_DESDE(1) = ""
        End If
        
        If Not IsNull(!NUMERO_HASTA_1) Then
            If IsNumeric(!NUMERO_HASTA_1) Then
                NUMERO_HASTA(1) = !NUMERO_HASTA_1
                LETRA_HASTA(1) = ""
            Else
                NUMERO_HASTA(1) = ""
                LETRA_HASTA(1) = !NUMERO_HASTA_1
            End If
        Else
            NUMERO_HASTA(1) = ""
            LETRA_HASTA(1) = ""
        End If
        hojaEx.Cells(R, C_Caja) = Caja
        hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), PasoFinal & Caja & "\" & NombreArchivo
        hojaEx.Cells(R, C_Indice) = Indice(1)
        hojaEx.Cells(R, C_Descripcion) = UCase(Trim(Replace(Descripcion(1), ".", " ")))

       hojaEx.Cells(R, C_Fecha_desde).NumberFormat = "@"
       hojaEx.Cells(R, C_Fecha_desde) = FECHA_DESDE(1)
       hojaEx.Cells(R, C_Fecha_hasta).NumberFormat = "@"

       
       
       
        If FECHA_HASTA(1) = "" And FECHA_DESDE(1) <> "" Then
            Rem hojaEx.Cells(R, C_Fecha_hasta) = hojaEx.Cells(R, C_Fecha_desde
             hojaEx.Cells(R, C_Fecha_hasta) = Format(hojaEx.Cells(R, C_Fecha_desde), "DD/MM/YYYY")
        Else
            Rem hojaEx.Cells(R, C_Fecha_hasta) = " " & FECHA_HASTA(1) & " "
            hojaEx.Cells(R, C_Fecha_hasta).value = FECHA_HASTA(1)
        End If
        hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(1)
        hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(1)
        hojaEx.Cells(R, C_Letra_Desde) = LETRA_DESDE(1)
        hojaEx.Cells(R, C_Letra_Hasta) = LETRA_HASTA(1)
        hojaEx.Cells(R, C_TIPO) = TIPO_ENVIO
        hojaEx.Cells(R, C_CSID) = !CSID
        R = R + 1
   
        '___________________________ FIN 1 ___________________________
            
    
                               '_______________________________ INICIO 2________________

                                 If Not IsNull(!IDEM_INDICE_2) Then
                                     Indice(2) = Indice(1)
                                 Else
                                     If Not IsNull(!INDICE_2) Then
                                            If IsNumeric(!INDICE_2) Then
                                                Indice(2) = !INDICE_2
                                             Else
                                                Indice(2) = "0"
                                            End If
                                         Else
                                            Indice(2) = "0"
                                         End If
                                 End If

                                Rem  Debug.Assert Caja <> 970629
                                 If IsNull(!IDEM_DETALLE_2) Then
                                     If Not IsNull(!DESCRIPCION_2) Then
                                            Descripcion(2) = !DESCRIPCION_2
                                            If !Form_Id = 189 And IsNumeric(Descripcion(2)) Then
                                                Descripcion(2) = Indice_Secundario_Disco(CInt(Descripcion(2)))
                                            End If
                                      Else
                                        Descripcion(2) = ""
                                     End If
                                 Else
                                    Descripcion(2) = Trim(Descripcion(1))
                                  End If

                                FECHA_DESDE(2) = FECHA_DESDE_SOLUCION(!DIA_DESDE_2, !MES_DESDE_2, !AÑO_DESDE_2)
                                FECHA_HASTA(2) = FECHA_HASTA_SOLUCION(!DIA_HASTA_2, !MES_HASTA_2, !AÑO_HASTA_2)
                                 
                                 If Not IsNull(!NUMERO_DESDE_2) Then
                                     If IsNumeric(!NUMERO_DESDE_2) Then
                                         NUMERO_DESDE(2) = !NUMERO_DESDE_2
                                         LETRA_DESDE(2) = ""
                                     Else
                                         NUMERO_DESDE(2) = ""
                                         LETRA_DESDE(2) = !NUMERO_DESDE_2
                                     End If
                                 Else
                                     NUMERO_DESDE(2) = ""
                                     LETRA_DESDE(2) = ""
                                 End If

                                 If Not IsNull(!NUMERO_HASTA_2) Then
                                     If IsNumeric(!NUMERO_HASTA_2) Then
                                         NUMERO_HASTA(2) = !NUMERO_HASTA_2
                                         LETRA_HASTA(2) = ""
                                     Else
                                         NUMERO_HASTA(2) = ""
                                         LETRA_HASTA(2) = !NUMERO_HASTA_2
                                     End If
                                 Else
                                     NUMERO_HASTA(2) = ""
                                     LETRA_HASTA(2) = ""
                                 End If

                                 Rem Control Luis
                                 If Not (Indice(2) = "0" And Descripcion(2) = "" And FECHA_DESDE(2) = "" And NUMERO_DESDE(2) = "") Then

                                     hojaEx.Cells(R, C_Caja) = Caja
                                     hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), PasoFinal & Caja & "\" & NombreArchivo
                                     hojaEx.Cells(R, C_Indice) = Indice(2)
                                     hojaEx.Cells(R, C_Descripcion) = UCase(Trim(Replace(Descripcion(2), ".", " ")))
                                     Rem hojaEx.Cells(R, C_Fecha_desde) = " " & FECHA_DESDE(2) & " "
                                     hojaEx.Cells(R, C_Fecha_desde).NumberFormat = "@"
                                     hojaEx.Cells(R, C_Fecha_desde) = Format(FECHA_DESDE(2), "DD/MM/YYYY")
                                    hojaEx.Cells(R, C_Fecha_hasta).NumberFormat = "@"

                                     If FECHA_HASTA(2) = "" And FECHA_DESDE(2) <> "" Then
                                      hojaEx.Cells(R, C_Fecha_hasta) = hojaEx.Cells(R, C_Fecha_desde)
                                     Else
                                         hojaEx.Cells(R, C_Fecha_hasta) = FECHA_HASTA(2)
                                     End If
                                     hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(2)
                                     hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(2)
                                     hojaEx.Cells(R, C_Letra_Desde) = LETRA_DESDE(2)
                                     hojaEx.Cells(R, C_Letra_Hasta) = LETRA_HASTA(2)
                                     hojaEx.Cells(R, C_TIPO) = TIPO_ENVIO
                                     hojaEx.Cells(R, C_CSID) = !CSID
                                     R = R + 1
                                 End If


                                 '___________________________ FIN 2 ___________________________




                               '_______________________ INICIO 3________________

                                 If Not IsNull(!IDEM_INDICE_3) Then
                                     Indice(3) = Indice(2)
                                 Else
                                     If Not IsNull(!INDICE_3) Then
                                            If IsNumeric(!INDICE_3) Then
                                                Indice(3) = !INDICE_3
                                             Else
                                                Indice(3) = "0"
                                            End If
                                         Else
                                            Indice(3) = "0"
                                         End If
                                 End If


                                 If IsNull(!IDEM_DETALLE_3) Then
                                     If Not IsNull(!DESCRIPCION_3) Then
                                          Descripcion(3) = Trim(Replace(!DESCRIPCION_3, ".", " "))
                                           If !Form_Id = 189 And IsNumeric(Descripcion(3)) Then
                                                Descripcion(3) = Indice_Secundario_Disco(CInt(Descripcion(3)))
                                           End If
                                     Else
                                        Descripcion(3) = ""
                                     End If
                                 Else
                                    Descripcion(3) = Trim(Descripcion(2))
                                  End If

                                FECHA_DESDE(3) = FECHA_DESDE_SOLUCION(!DIA_DESDE_3, !MES_DESDE_3, !AÑO_DESDE_3)
                                FECHA_HASTA(3) = FECHA_HASTA_SOLUCION(!DIA_HASTA_3, !MES_HASTA_3, !AÑO_HASTA_3)


                                 If Not IsNull(!NUMERO_DESDE_3) Then
                                     If IsNumeric(!NUMERO_DESDE_3) Then
                                         NUMERO_DESDE(3) = !NUMERO_DESDE_3
                                         LETRA_DESDE(3) = ""
                                     Else
                                         NUMERO_DESDE(3) = ""
                                         LETRA_DESDE(3) = !NUMERO_DESDE_3
                                     End If
                                 Else
                                     NUMERO_DESDE(3) = ""
                                     LETRA_DESDE(3) = ""
                                 End If

                                 If Not IsNull(!NUMERO_HASTA_3) Then
                                     If IsNumeric(!NUMERO_HASTA_3) Then
                                         NUMERO_HASTA(3) = !NUMERO_HASTA_3
                                         LETRA_HASTA(3) = ""
                                     Else
                                         NUMERO_HASTA(3) = ""
                                         LETRA_HASTA(3) = !NUMERO_HASTA_3
                                     End If
                                 Else
                                     NUMERO_HASTA(3) = ""
                                     LETRA_HASTA(3) = ""
                                 End If

                                Rem Control Luis
                                 If Not (Indice(3) = "0" And Descripcion(3) = "" And FECHA_DESDE(3) = "" And NUMERO_DESDE(3) = "") Then

                                     hojaEx.Cells(R, C_Caja) = Caja
                                     hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), PasoFinal & Caja & "\" & NombreArchivo
                                     hojaEx.Cells(R, C_Indice) = Indice(3)
                                     hojaEx.Cells(R, C_Descripcion) = UCase(Trim(Replace(Descripcion(3), ".", " ")))
                                     Rem hojaEx.Cells(R, C_Fecha_desde) = " " & FECHA_DESDE(3) & " "
                                     hojaEx.Cells(R, C_Fecha_desde).NumberFormat = "@"
                                     hojaEx.Cells(R, C_Fecha_desde) = Format(FECHA_DESDE(3), "DD/MM/YYYY")
                                     hojaEx.Cells(R, C_Fecha_hasta).NumberFormat = "@"

                                     If FECHA_HASTA(3) = "" And FECHA_DESDE(3) <> "" Then
                                      hojaEx.Cells(R, C_Fecha_hasta) = hojaEx.Cells(R, C_Fecha_desde)
                                     Else
                                         hojaEx.Cells(R, C_Fecha_hasta) = FECHA_HASTA(3)
                                     End If
                                     hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(3)
                                     hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(3)
                                     hojaEx.Cells(R, C_Letra_Desde) = LETRA_DESDE(3)
                                     hojaEx.Cells(R, C_Letra_Hasta) = LETRA_HASTA(3)
                                     hojaEx.Cells(R, C_TIPO) = TIPO_ENVIO
                                     hojaEx.Cells(R, C_CSID) = !CSID
                                     R = R + 1
                                 End If


                                 '___________________________ FIN 3 ___________________________


                                  '_______________________ INICIO 4________________

                                 If Not IsNull(!IDEM_INDICE_4) Then
                                     Indice(4) = Indice(3)
                                 Else
                                     If Not IsNull(!INDICE_4) Then
                                            If IsNumeric(!INDICE_4) Then
                                                Indice(4) = !INDICE_4
                                             Else
                                                Indice(4) = "0"
                                            End If
                                         Else
                                            Indice(4) = "0"
                                         End If
                                 End If


                            If IsNull(!IDEM_DETALLE_4) Then
                                If Not IsNull(!DESCRIPCION_4) Then
                                        Descripcion(4) = Trim(Replace(!DESCRIPCION_4, ".", " "))
                                        If !Form_Id = 189 And IsNumeric(Descripcion(4)) Then
                                            Descripcion(4) = Indice_Secundario_Disco(CInt(Descripcion(4)))
                                        
                                        End If
                                 Else
                                 Descripcion(4) = ""
                                 End If
                            Else
                                Descripcion(4) = Trim(Descripcion(3))
                            End If


                                 FECHA_DESDE(4) = FECHA_DESDE_SOLUCION(!DIA_DESDE_4, !MES_DESDE_4, !AÑO_DESDE_4)
                                FECHA_HASTA(4) = FECHA_HASTA_SOLUCION(!DIA_HASTA_4, !MES_HASTA_4, !AÑO_HASTA_4)

                                 
                                 If Not IsNull(!NUMERO_DESDE_4) Then
                                     If IsNumeric(!NUMERO_DESDE_4) Then
                                         NUMERO_DESDE(4) = !NUMERO_DESDE_4
                                         LETRA_DESDE(4) = ""
                                     Else
                                         NUMERO_DESDE(4) = ""
                                         LETRA_DESDE(4) = !NUMERO_DESDE_4
                                     End If
                                 Else
                                     NUMERO_DESDE(4) = ""
                                     LETRA_DESDE(4) = ""
                                 End If

                                 If Not IsNull(!NUMERO_HASTA_4) Then
                                     If IsNumeric(!NUMERO_HASTA_4) Then
                                         NUMERO_HASTA(4) = !NUMERO_HASTA_4
                                         LETRA_HASTA(4) = ""
                                     Else
                                         NUMERO_HASTA(4) = ""
                                         LETRA_HASTA(4) = !NUMERO_HASTA_4
                                     End If
                                 Else
                                     NUMERO_HASTA(4) = ""
                                     LETRA_HASTA(4) = ""
                                 End If

                                Rem Control Luis
                                 If Not (Indice(4) = "0" And Descripcion(4) = "" And FECHA_DESDE(4) = "" And NUMERO_DESDE(4) = "") Then
                                     hojaEx.Cells(R, C_Caja) = Caja
                                     hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), PasoFinal & Caja & "\" & NombreArchivo
                                     hojaEx.Cells(R, C_Indice) = Indice(4)
                                     hojaEx.Cells(R, C_Descripcion) = UCase(Trim(Replace(Descripcion(4), ".", " ")))
                                     Rem hojaEx.Cells(R, C_Fecha_desde) = " " & FECHA_DESDE(4) & " "
                                     hojaEx.Cells(R, C_Fecha_desde).NumberFormat = "@"
                                     hojaEx.Cells(R, C_Fecha_desde) = Format(FECHA_DESDE(4), "DD/MM/YYYY")
                                     hojaEx.Cells(R, C_Fecha_hasta).NumberFormat = "@"

                                     If FECHA_HASTA(4) = "" And FECHA_DESDE(4) <> "" Then
                                         hojaEx.Cells(R, C_Fecha_hasta) = hojaEx.Cells(R, C_Fecha_desde)
                                     Else
                                         hojaEx.Cells(R, C_Fecha_hasta) = FECHA_HASTA(4)
                                     End If
                                     hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(4)
                                     hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(4)
                                     hojaEx.Cells(R, C_Letra_Desde) = LETRA_DESDE(4)
                                     hojaEx.Cells(R, C_Letra_Hasta) = LETRA_HASTA(4)
                                     hojaEx.Cells(R, C_TIPO) = TIPO_ENVIO
                                     hojaEx.Cells(R, C_CSID) = !CSID
                                     R = R + 1
                                 End If


                                 '___________________________ FIN 4 ___________________________

            .MoveNext
            
            Loop
          End With
         
         
         
         
          Dim Paso_planilla As String

        Paso_planilla = "\\222.15.19.251\basa\Administracion\Referencias\" & InputBox("Ingrese el nombre de la planilla") & "    " & Format(Now, "ddmmyyy hhss") & ".xls"

                libroEx.SaveAs Paso_planilla
                libroEx.Close
                ApExcel.Quit
                Set hojaEx = Nothing
                Set libroEx = Nothing
                Set ApExcel = Nothing

                MsgBox "Terminado"
            If MsgBox("Usted quiere ver la planilla", vbYesNo) = vbYes Then
            Shell "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE " & Chr(34) & Paso_planilla & Chr(34), vbNormalFocus
            End If
         
         
         
         
         
         
         
         
         
         
         
'           Debug.Assert RsPlanilla!CAJA_Nº <> 875770
'
'            If i > 4 Then
'                i = 1
'                Exit For
'            End If
'
'                For i = 1 To 4
'                    If Not IsNull(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))) Then
'                        Rem MsgBox RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))
'                        FECHA_DESDE(i) = Format(Format(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_DESDE_" & CStr(i)), "00"), "DD/MM/YYYY")
'                        If Not IsNull(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i))) Then
'
'
'                           If RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)) = 0 Then
'                                FECHA_HASTA(i) = FECHA_DESDE(i)
'                           Else
'
'                                If Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") <> "00" Then
'                                    FECHA_HASTA(i) = Format(Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_HASTA_" & CStr(i)), "00"), "DD/MM/YYYY")
'                                Else
'                                    FECHA_HASTA(i) = FECHA_DESDE(i)
'                                End If
'                            End If
'                        Else
'                            FECHA_HASTA(i) = FECHA_DESDE(i)
'                        End If
'                    Else
'                        FECHA_DESDE(i) = ""
'                    End If
'                    If Not IsNull(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))) Then
'                        Rem MsgBox RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))
'                        NUMERO_DESDE(i) = Trim(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i)))
'                        If Format(RsPlanilla.Fields.Item("NUMERO_HASTA_" & CStr(i)), "") <> "" Then
'                            NUMERO_HASTA(i) = Trim(RsPlanilla.Fields.Item("NUMERO_HASTA_" & i))
'                        Else
'                             NUMERO_HASTA(i) = NUMERO_DESDE(i)
'                        End If
'                    Else
'                        NUMERO_DESDE(i) = ""
'                        NUMERO_HASTA(i) = ""
'                    End If
'
'                    If Not IsNull(RsPlanilla.Fields.Item("INDICE_" & CStr(i))) Then
'                        INDICE(i) = RsPlanilla.Fields.Item("INDICE_" & CStr(i))
'                        Else
'                            If i <> 1 Then
'                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_INDICE_" & CStr(i))) Then
'                                    INDICE(i) = INDICE(CStr(i - 1))
'                                 Else
'                                    If Not IsNull(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))) Or Not IsNull(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))) Or Not IsNull(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))) Then
'                                     INDICE(i) = "0"
'                                    Else
'                                     INDICE(i) = ""
'                                    End If
'                                End If
'                            Else
'                                If INDICE(i) = "" Then
'                                INDICE(i) = "0"
'                                Else
'                                End If
'
'                            End If
'
'                    End If
'
'                    If Not IsNull(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))) Then
'                        Descripcion(i) = LCase(Trim(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))))
'                    Else
'                            If i <> 1 Then
'                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_DETALLE_" & CStr(i))) Then
'                                If RsPlanilla.Fields.Item("IDEM_DETALLE_" & CStr(i)) <> 0 Then
'
'                                    Descripcion(i) = LCase(Descripcion(CStr(i - 1)))
'                                 Else
'                                 Descripcion(i) = ""
'                                 End If
'                                Else
'                                Descripcion(i) = ""
'                                End If
'                            Else
'                                Descripcion(i) = ""
'                            End If
'                    End If
'
'
'
'
'
'                If IsNull(RsPlanilla.Fields.Item("CAJA_Nº").value) Then
'                        directorio = "Error"
'                        hojaEx.Cells(R, C_Caja) = "Error"
'                        hojaEx.Cells(R, C_Indice) = ""
'                        hojaEx.Cells(R, C_Descripcion) = ""
'                        hojaEx.Cells(R, C_Fecha_desde) = ""
'                        hojaEx.Cells(R, C_Fecha_hasta) = ""
'                        hojaEx.Cells(R, C_N°_Desde) = ""
'                        hojaEx.Cells(R, C_N°_Hasta) = ""
'                        GoTo ERRORCAJA:
'                Else
'                        NombreArchivo = RsPlanilla.Fields.Item("CAJA_Nº").value & "_" & FechaHora & ".tif"
'               End If
'               Select Case RsPlanilla!ENVIO_CAJAS
'                        Case 1
'                            hojaEx.Cells(R, C_TIPO) = "Sin Refrencia"
'                            Descripcion(i) = Descripcion(i) & " //" & hojaEx.Cells(R, C_TIPO)
'                        Case 2
'                            hojaEx.Cells(R, C_TIPO) = "Envio Por correo"
'                            Descripcion(i) = Descripcion(i) & " // " & hojaEx.Cells(R, C_TIPO)
'                        Case 3
'                            hojaEx.Cells(R, C_TIPO) = "Referencia en planta"
'                            Descripcion(i) = Descripcion(i) & " // " & hojaEx.Cells(R, C_TIPO)
'                        Case 4
'                            hojaEx.Cells(R, C_TIPO) = "Cargar Legajos"
'                            Descripcion(i) = Descripcion(i) & " // " & hojaEx.Cells(R, C_TIPO)
'                        Case 5
'                            hojaEx.Cells(R, C_TIPO) = "Digitalizar"
'                            Descripcion(i) = Descripcion(i) & " " & hojaEx.Cells(R, C_TIPO)
'                        End Select
'
'
'              If INDICE(i) = "" And Trim(RsPlanilla.Fields.Item("CAJA_Nº").value) <> "" And hojaEx.Cells(R, C_TIPO) <> "" Then
'                    ControlIndice = True
'                    INDICE(i) = "1"
'                    Descripcion(1) = Descripcion(1) & " " & hojaEx.Cells(R, C_TIPO)
'              Else
'                  ControlIndice = False
'              End If
'
'
'                    If Trim(INDICE(i)) <> "" Then
'                        hojaEx.Cells(R, C_Caja) = RsPlanilla.Fields.Item("CAJA_Nº").value
'
'                        hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & Trim(RsPlanilla.Fields.Item("CAJA_Nº").value) & "\" & NombreArchivo
'                        hojaEx.Cells(R, C_Indice) = INDICE(i)
'                        Rem Debug.Print "Incide: " & Indice(i)
'                        hojaEx.Cells(R, C_Descripcion) = Trim(Replace(Replace(Descripcion(i), vbCrLf, " "), vbCr, " "))
'
'                        If Trim(CStr(FECHA_DESDE(i))) <> "" Then
'                            If IsDate(FECHA_DESDE(i)) Then
'
'
'                                hojaEx.Cells(R, C_Fecha_desde).value = " " & CStr(FECHA_DESDE(i))
'
'                            Else
'                                hojaEx.Cells(R, C_Fecha_desde) = ""
'                            End If
'
'                        Else
'                            hojaEx.Cells(R, C_Fecha_desde).value = Format(FECHA_DESDE(i), "DD/MM/YYYY")
'                        End If
'                       Debug.Print "Caja : " & RsPlanilla.Fields.Item("CAJA_Nº").value & "//Incide: " & INDICE(i) & "//Fecha : " & CStr(FECHA_DESDE(i)) & "//dese:" & hojaEx.Cells(R, C_Descripcion) & " //Numero :" & NUMERO_DESDE(i)
'                       Rem  Debug.Print "Fecha : " & CStr(FECHA_DESDE(i))
'
'                       If Trim(CStr(FECHA_DESDE(i))) <> "" Then
'                            If IsDate(Format(FECHA_HASTA(i), "DD/MM/YYYY")) Then
'                                hojaEx.Cells(R, C_Fecha_hasta).value = " " & Format(FECHA_HASTA(i), "DD/MM/YYYY")
'                            Else
'                                hojaEx.Cells(R, C_Fecha_hasta).value = ""
'                            End If
'
'                       Else
'                            hojaEx.Cells(R, C_Fecha_hasta) = ""
'                       End If
'                       If Trim(NUMERO_DESDE(i)) <> "" Then
'                        hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(i)
'                        hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(i)
'                       Else
'                            hojaEx.Cells(R, C_N°_Desde) = ""
'                            hojaEx.Cells(R, C_N°_Hasta) = ""
'                            NUMERO_DESDE(i) = ""
'                            NUMERO_HASTA(i) = ""
'                       End If
'
'              If ControlIndice = True Then
'                 i = 5
'                ControlIndice = False
'              End If
'
'ERRORCAJA:
'
'
'
'                          hojaEx.Cells(R, C_CSID) = RsPlanilla.Fields.Item("CSID").value
'                          lblCantidadRegistros.Caption = RsPlanilla.Fields.Item("CSID").value
'
'                        R = R + 1
'                    End If
'                Next
'
'
'                If Dir("\\222.15.19.222\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3)) <> "" Then
'
'                If Not IsNull(RsPlanilla.Fields.Item("CAJA_Nº")) Then
'                       If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\ " & CStr(RsPlanilla.Fields.Item("CAJA_Nº")), vbDirectory) = "" Then
'                            MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & CStr(RsPlanilla.Fields.Item("CAJA_Nº").value)
'                            FileCopy "\\222.15.19.222\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
'                         Else
'                            FileCopy "\\222.15.19.222\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
'                        End If
'
'                        Else
'                            Rem MsgBox "No se encontro La imagen " & RsPlanilla.Fields.Item("Suspense_File").value
'                        End If
'                  Else
'                        If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\", vbDirectory) = "" Then
'                            MkDir ("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\")
'                            FileCopy "\\TELEFORM\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\" & RsPlanilla.Index & Format(Now, "ddmmyyyy mmss") & ".tif"
'                         Else
'                            FileCopy "\" & Mid(RsPlanilla.Fields.Item("Suspense_File").value, 3), "Z:\Referencias" & "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Cajas\Error\" & RsPlanilla.Index & Format(Now, "ddmmyyyy mmss") & ".tif"
'                        End If
'                   End If
'
'
'Proximo:
'
                 
               '  RsPlanilla.MoveNext
'        Loop
'
'        Dim Paso_planilla As String
'
'        Paso_planilla = "\\222.15.19.251\basa\Administracion\Referencias\" & InputBox("Ingrese el nombre de la planilla") & "    " & Format(Now, "ddmmyyy hhss") & ".xls"
'
'                libroEx.SaveAs Paso_planilla
'                libroEx.Close
'                ApExcel.Quit
'                Set hojaEx = Nothing
'                Set libroEx = Nothing
'                Set ApExcel = Nothing
'
'                MsgBox "Terminado"
'            If MsgBox("Usted quiere ver la planilla", vbYesNo) = vbYes Then
'            Shell "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE " & Chr(34) & Paso_planilla & Chr(34), vbNormalFocus
'            End If


End Sub
Private Sub cmdReferenciasCargadasControl_Click()
 Dim Paso As String
        CommonDialog1.FileName = "\\Serverbackup_1\E\Usuarios\Clientes\*.xls"
        CommonDialog1.ShowOpen
        Paso = CommonDialog1.FileName
        If CommonDialog1.FileTitle = "*.xls" Then
        Exit Sub
        End If
    If Control_Excel_Referencias_Cargadas(Paso, True) Then
        MsgBox "Existe un errores", vbCritical, "Control Carga"
      Else
       MsgBox "NO Existe un errores", vbInformation, "Control Carga"
    End If
End Sub

Private Sub cmdRecibos_Click()
Dim comRecibos As New ADODB.Connection
            Dim RsRecibos As New ADODB.Recordset
            Dim Sql As String
            Dim Sql_Filtro As String
            Dim Imagen As String
            Dim Directorio_Remito As String
            Dim Remito_Manual As String
            Dim Remito_Manual_Error As String
            Dim registro As Integer
            Dim intimagen As Integer
            
            
            comRecibos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=Z:\Sistemas\Basa\BaseTeleform\Referencias.mdb"
            
            
          Sql = "   SELECT Suspense_File, BatchNo, BatchPgNo,ID "
          Sql = Sql & " ,  Numero, Fecha, Nombre , PASADO "
          Sql = Sql & "  FROM RECIBOBASA"
             Sql = Sql & " WHERE PASADO=False"

          Sql = Sql & " ORDER BY BatchNo,BatchPgNo "
           RsRecibos.Open Sql, comRecibos
           
          Dim fecha As String
          Dim recibo As String
          Dim Nombre As String

Do While Not RsRecibos.EOF
lblCantidadRegistros.Caption = CStr(RsRecibos!BatchNo) & "-" & CStr(RsRecibos!BatchPgNo)
lblCantidadRegistros.Refresh

If Not IsNull(RsRecibos!fecha) Then
    fecha = Trim(Replace(Replace(RsRecibos!fecha, "/", "_"), "*", " "))
Else
 fecha = "00_00_00"
End If

If Not IsNull(RsRecibos!NUMERO) Then
recibo = Trim(Replace(RsRecibos!NUMERO, "/", "_"))
 
Else
 recibo = 0
End If

If Not IsNull(RsRecibos!Nombre) Then
    Nombre = Replace(Replace(Replace(Trim(UCase(RsRecibos!Nombre)), "-", "_"), "/", "_"), "*", " ")
Else
    Nombre = "XXXXXX"
End If




FileCopy "\\TELEFORM" & Mid(RsRecibos!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Recibos\" & Nombre & "  " & fecha & "  Recibo " & recibo & "   ID " & RsRecibos!ID & ".tif"
    




RsRecibos.MoveNext

Loop

MsgBox "Terminado"



End Sub

Private Sub cmdRemitos_Click()
            Dim comRemitos As New ADODB.Connection
            Dim RsRemitos As New ADODB.Recordset
            Dim Sql As String
            Dim Sql_Filtro As String
            Dim Imagen As String
            Dim Directorio_Remito As String
            Dim Remito_Manual As String
            Dim Remito_Manual_Error As String
            Dim registro As Integer
            Dim intimagen As Integer
            
            
           
            
            
          Sql = "   SELECT Suspense_File, BatchNo, BatchPgNo,BatchPgDta, ID "
          Sql = Sql & " , NumeroRemito_1, REMITO_MANUAL, NumeroRemito,ENBASA "
          Sql = Sql & " From TELEFORM_REMITO "
           Sql = Sql & " WHERE  (NOT (BatchPgDta IS NULL)) and ENBASA IS NULL"
         Rem Sql = Sql & " where (ID >  23442)"
          Sql = Sql & " ORDER BY BatchNo,BatchPgNo "
          RsRemitos.Open Sql, strConBasa
           
            

Do While Not RsRemitos.EOF
lblCantidadRegistros.Caption = CStr(RsRemitos!BatchNo) & "-" & CStr(RsRemitos!BatchPgNo)
lblCantidadRegistros.Refresh
Sql_Filtro = ""
If Not IsNull(RsRemitos!Remito_Manual) Then
    If Len(RsRemitos!Remito_Manual) = 13 Then
        Remito_Manual = Trim(RsRemitos!Remito_Manual)
    Else
        Remito_Manual = Format(Trim(RsRemitos!Remito_Manual), "0001-00000000")
    End If
    Sql_Filtro = " NRO_REM_PROV = '" & Remito_Manual & "'"
    Imagen = Remito_Manual & "_" & Format(Now, "ddMMyy_mmss") & ".tif"
    
   
    
    Directorio_Remito = Remito_Manual
End If

If Not IsNull(RsRemitos!NumeroRemito) Then
    Sql_Filtro = " NRO_REMITO = " & CLng(RsRemitos!NumeroRemito)
    Imagen = Trim(RsRemitos!NumeroRemito) & "_" & Format(Now, "mmss") & ".tif"
    Directorio_Remito = CStr(CLng(RsRemitos!NumeroRemito))
    Sql = " Update REQUERIMIENTO "
    Sql = Sql & vbCrLf & " Set IDESTADO = 8 "
    Sql = Sql & vbCrLf & " Where  IDREMITO = " & CLng(RsRemitos!NumeroRemito)
    ExecutarSql (Sql)
End If



  If Sql_Filtro <> "" Then
    Sql = " Update REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " SET IMAGEN = '" & Imagen & "'"
    Sql = Sql & vbCrLf & " WHERE   "
    Sql = Sql & Sql_Filtro
    registro = ExecutarSql(Sql)
Else
registro = 0
End If



If registro = 0 Then
        Remito_Manual_Error = Trim(Remito_Manual_Error) & RsRemitos!NumeroRemito & Remito_Manual & vbCrLf
        intimagen = intimagen + 1
        
       
        
        If Dir("\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3)) = "" Then
            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 0"
            Sql = Sql & " WHERE ID = " & RsRemitos!ID
            ExecutarSql Sql
        
        Else
        
        
        
            FileCopy "\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\error\" & CStr(intimagen) & ".TIF"
            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 1"
            Sql = Sql & " WHERE ID = " & RsRemitos!ID
            ExecutarSql Sql
        End If
        
    Else
    
     AcutalizarImagenRemito Mid(RsRemitos!BatchPgDta, 1, Len(RsRemitos!BatchPgDta) - 9), Sql_Filtro
        If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Directorio_Remito, vbDirectory) = "" Then
       Rem  FileSystem.MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Directorio_Remito
        End If
        
        If Dir("\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3)) = "" Then
            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 0"
            Sql = Sql & " WHERE ID = " & RsRemitos!ID
            ExecutarSql Sql
        Else
            Rem FileCopy "\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Directorio_Remito & "\" & Imagen
            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 1"
            Sql = Sql & " WHERE ID = " & RsRemitos!ID
            ExecutarSql Sql
        End If


  
    
End If





RsRemitos.MoveNext

Loop

Clipboard.Clear
Clipboard.SetText Remito_Manual_Error
MsgBox "Terminado"



End Sub

Private Sub cmdRemitosNuevos_Click()



    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim DatoRemito As String
    
    
    
        Sql = " SELECT DOCUMENTOS_DIGITALES.DIRECTORIO_PASO ,DOCUMENTOS_DIGITALES.ID , DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, "
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, DOCUMENTOS_DIGITALES_LOTE.FECHA_INGRESO, DOCUMENTOS_DIGITALES.LETRA_DESDE,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.LETRA_HASTA , DOCUMENTOS_DIGITALES_LOTE.LOTE_ESTADO, DOCUMENTOS_DIGITALES.estado"
        Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 83) AND (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES in( 3802, 10984))  AND"
        Sql = Sql & vbCrLf & " (DOCUMENTOS_DIGITALES_LOTE.FECHA_INGRESO > CONVERT(DATETIME, '2017-05-30 00:00:00', 102)) AND"
        Sql = Sql & vbCrLf & " (NOT (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)) AND  ESTADO  = 'LISTA PARA EXPORTAR' "
        Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.LETRA_DESDE"

rs.CursorLocation = adUseClient

rs.Open Sql, strConBasa, adOpenForwardOnly, adLockReadOnly

Dim Imagen As String
Dim Registros As Integer
Dim rsCliente As New ADODB.Recordset



Do While Not rs.EOF
    
    Registros = 0
    If Not IsNull(rs!LETRA_DESDE) Then
    
    
            DatoRemito = Trim(rs!LETRA_DESDE)
            Debug.Print DatoRemito
            If Mid(DatoRemito, 1, 4) = "0001" Then
                If DatoRemito <> "0001-000_____" Then
                    
                    
                    
                    If Len(DatoRemito) = 14 Then
                        DatoRemito = "0001-00" & Mid(DatoRemito, 9)
                    End If
                     
                    
                     
                     Imagen = DatoRemito & "_" & rs!ID & ".tif"
                     Sql = " Update REMITOS_CUERPO "
                     Sql = Sql & vbCrLf & " SET IMAGEN = '" & Imagen & "'"
                     Sql = Sql & vbCrLf & " WHERE   "
                     Sql = Sql & vbCrLf & "  NRO_REMITO > 179501 and   NRO_REM_PROV = '" & DatoRemito & "'"
                     Registros = ExecutarSql(Sql)
                     If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & DatoRemito, vbDirectory) = "" Then
                         FileSystem.MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & DatoRemito
                     End If
                     FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & DatoRemito & "\" & Imagen
                    
                    
                    If Registros > 0 Then
                        Set rsCliente = New ADODB.Recordset
                        Sql = " SELECT CLIENTES.RAZON_SOCIAL"
                        Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO INNER JOIN"
                        Sql = Sql & vbCrLf & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
                        Sql = Sql & vbCrLf & " WHERE NRO_REMITO > 179501 and  REMITOS_CUERPO.NRO_REM_PROV = '" & DatoRemito & "'"
                        
                        rsCliente.Open Sql, strConBasa
                        If Not rsCliente.EOF Then
                            Sql = " Update DOCUMENTOS_DIGITALES "
                            Sql = Sql & vbCrLf & " SET LETRA_HASTA ='" & Trim(rsCliente!RAZON_SOCIAL) & "'"
                            Sql = Sql & vbCrLf & " Where LETRA_HASTA IS NULL  AND ID = " & rs!ID
                            Rem Sql = Sql & vbCrLf & " Where  ID = " & rs!ID
                            ExecutarSql Sql
                        End If
                    Else
                        Sql = " Update DOCUMENTOS_DIGITALES "
                        Sql = Sql & vbCrLf & " SET LETRA_HASTA ='SIN CLIENTE'"
                        Sql = Sql & vbCrLf & " Where LETRA_HASTA IS NULL  AND ID = " & rs!ID
                        ExecutarSql Sql
                        
                    End If
                    
                    
                    Rem 'cc') AND (REMITOS_CUERPO.NRO_REMITO = 11)
                    
                
                End If
            End If
                
            If IsNumeric(DatoRemito) Then
                If CLng(DatoRemito) > 179758 Then
                    Imagen = DatoRemito & "_" & rs!ID & ".tif"
                    
                     Sql = " Update REQUERIMIENTO "
                    Sql = Sql & vbCrLf & " Set IDESTADO = 8 "
                    Sql = Sql & vbCrLf & " Where IDESTADO > 5 and  IDREMITO = " & DatoRemito
                    Registros = ExecutarSql(Sql)
                    If Registros = 0 Then
                        Debug.Print DatoRemito
                    End If
                    
                    
                    Sql = " Update REMITOS_CUERPO "
                    Sql = Sql & vbCrLf & " SET IMAGEN = '" & Imagen & "'"
                    Sql = Sql & vbCrLf & " WHERE   "
                    Sql = Sql & vbCrLf & " NRO_REMITO = " & DatoRemito
                    Registros = ExecutarSql(Sql)
                    
                    If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Trim(rs!LETRA_DESDE), vbDirectory) = "" Then
                        FileSystem.MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Trim(rs!LETRA_DESDE)
                    End If
                    FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & DatoRemito & "\" & Imagen
                
                
                    If Registros > 0 Then
                        Set rsCliente = New ADODB.Recordset
                        Sql = " SELECT CLIENTES.RAZON_SOCIAL"
                        Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO INNER JOIN"
                        Sql = Sql & vbCrLf & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
                        Sql = Sql & vbCrLf & " WHERE NRO_REMITO = " & DatoRemito
                        rsCliente.Open Sql, strConBasa
                        If Not rsCliente.EOF Then
                            Sql = " Update DOCUMENTOS_DIGITALES "
                            Sql = Sql & vbCrLf & " SET LETRA_HASTA ='" & Trim(rsCliente!RAZON_SOCIAL) & "'"
                            Sql = Sql & vbCrLf & " Where  LETRA_HASTA IS NULL  AND ID = " & rs!ID
                            ExecutarSql Sql
                        End If
                    Else
                        Sql = " Update DOCUMENTOS_DIGITALES "
                        Sql = Sql & vbCrLf & " SET LETRA_HASTA ='SIN CLIENTE'"
                        Sql = Sql & vbCrLf & " Where  LETRA_HASTA IS NULL  AND ID = " & rs!ID
                        ExecutarSql Sql
                        
                    End If
                
                End If
            End If
            
            If Mid(DatoRemito, 1, 3) = "D10" Then
                DatoRemito = Mid(DatoRemito, 5)
                Imagen = DatoRemito & "_" & rs!ID & ".tif"
                Sql = " Update REQUERIMIENTO "
                Sql = Sql & vbCrLf & " Set IDESTADO = 8 "
                Sql = Sql & vbCrLf & " Where IDESTADO > 5 and IDREQUERIMIENTO = " & DatoRemito
                Registros = ExecutarSql(Sql)
                If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & DatoRemito, vbDirectory) = "" Then
                    FileSystem.MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & DatoRemito
                End If
                FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & DatoRemito & "\" & Imagen
            End If
            
            If Registros <> 0 Then
            Sql = " Update DOCUMENTOS_DIGITALES"
            Sql = Sql & vbCrLf & " SET ESTADO ='EXPORTADO'"
            Sql = Sql & vbCrLf & " Where ID = " & rs!ID
            ExecutarSql Sql
            End If
            
    
    End If
    
    rs.MoveNext
Loop

    
    
'    lblCantidadRegistros.Caption = CStr(RsRemitos!BatchNo) & "-" & CStr(RsRemitos!BatchPgNo)
'lblCantidadRegistros.Refresh
'Sql_Filtro = ""
'If Not IsNull(RsRemitos!Remito_Manual) Then
'    If Len(RsRemitos!Remito_Manual) = 13 Then
'        Remito_Manual = Trim(RsRemitos!Remito_Manual)
'    Else
'        Remito_Manual = Format(Trim(RsRemitos!Remito_Manual), "0001-00000000")
'    End If
'    Sql_Filtro = " NRO_REM_PROV = '" & Remito_Manual & "'"
'    Imagen = Remito_Manual & "_" & Format(Now, "ddMMyy_mmss") & ".tif"
'
'
'
'    Directorio_Remito = Remito_Manual
'End If
'
'If Not IsNull(RsRemitos!NumeroRemito) Then
'    Sql_Filtro = " NRO_REMITO = " & CLng(RsRemitos!NumeroRemito)
'    Imagen = Trim(RsRemitos!NumeroRemito) & "_" & Format(Now, "mmss") & ".tif"
'    Directorio_Remito = CStr(CLng(RsRemitos!NumeroRemito))
'    Sql = " Update REQUERIMIENTO "
'    Sql = Sql & vbCrLf & " Set IDESTADO = 8 "
'    Sql = Sql & vbCrLf & " Where  IDREMITO = " & CLng(RsRemitos!NumeroRemito)
'    ExecutarSql (Sql)
'End If
'
'
'
'  If Sql_Filtro <> "" Then
'    Sql = " Update REMITOS_CUERPO "
'    Sql = Sql & vbCrLf & " SET IMAGEN = '" & Imagen & "'"
'    Sql = Sql & vbCrLf & " WHERE   "
'    Sql = Sql & Sql_Filtro
'    registro = ExecutarSql(Sql)
'Else
'registro = 0
'End If
'
'
'
'If registro = 0 Then
'        Remito_Manual_Error = Trim(Remito_Manual_Error) & RsRemitos!NumeroRemito & Remito_Manual & vbCrLf
'        intimagen = intimagen + 1
'
'
'
'        If Dir("\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3)) = "" Then
'            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 0"
'            Sql = Sql & " WHERE ID = " & RsRemitos!ID
'            ExecutarSql Sql
'
'        Else
'
'
'
'            FileCopy "\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\error\" & CStr(intimagen) & ".TIF"
'            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 1"
'            Sql = Sql & " WHERE ID = " & RsRemitos!ID
'            ExecutarSql Sql
'        End If
'
'    Else
'
'     AcutalizarImagenRemito Mid(RsRemitos!BatchPgDta, 1, Len(RsRemitos!BatchPgDta) - 9), Sql_Filtro
'        If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Directorio_Remito, vbDirectory) = "" Then
'       Rem  FileSystem.MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Directorio_Remito
'        End If
'
'        If Dir("\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3)) = "" Then
'            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 0"
'            Sql = Sql & " WHERE ID = " & RsRemitos!ID
'            ExecutarSql Sql
'        Else
'            Rem FileCopy "\\PCTELEMEMO1" & Mid(RsRemitos!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Directorio_Remito & "\" & Imagen
'            Sql = " UPDATE TELEFORM_REMITO SET ENBASA = 1"
'            Sql = Sql & " WHERE ID = " & RsRemitos!ID
'            ExecutarSql Sql
'        End If
'
'
'
'
'End If
'
'
'
'
'
'RsRemitos.MoveNext
'
'Loop
'
'Clipboard.Clear
'Clipboard.SetText Remito_Manual_Error
'MsgBox "Terminado"
'
'
'
'
'
'
'
'
    





End Sub

Private Sub cmdRequerimiento_Click()

            Dim comRequer As New ADODB.Connection
            Dim RscomRequer As New ADODB.Recordset
            Dim Sql As String
            Dim Sql_Filtro As String
            Dim Imagen As String
            Dim Directorio_Remito As String
            Dim Remito_Manual As String
            Dim Remito_Manual_Error As String
            Dim registro As Integer
            Dim intimagen As Integer


            


          
            Sql = " SELECT *  From TELEFORM_REQUERIMIENTO  "
            Sql = Sql & " WHERE(ENBASA IS NULL) "
            Sql = Sql & " ORDER BY BatchNo,BatchPgNo "
          
           RscomRequer.Open Sql, strConBasa



Do While Not RscomRequer.EOF
                lblCantidadRegistros.Caption = CStr(RscomRequer!BatchNo) & "-" & CStr(RscomRequer!BatchPgNo)
                lblCantidadRegistros.Refresh
                If Not IsNull(RscomRequer!REQUERIMIENTO) Then
                    Directorio_Remito = RscomRequer!REQUERIMIENTO
                    Imagen = RscomRequer!REQUERIMIENTO & "  " & Format(Now, "ddMMyy_mmss") & ".tif"
                    Sql = " Update REQUERIMIENTO "
                    Sql = Sql & vbCrLf & " Set IDESTADO = 8 "
                    Sql = Sql & vbCrLf & " , imagen =  '" & Imagen & "'"
                    Sql = Sql & vbCrLf & " Where IDREQUERIMIENTO = " & RscomRequer!REQUERIMIENTO
                    registro = ExecutarSql(Sql)
                    
                 Rem   "\\\\222.15.19.130"
                    
                    If registro = 0 Then
                        Remito_Manual_Error = Trim(Remito_Manual_Error) & RscomRequer!REQUERIMIENTO & Remito_Manual & vbCrLf
                        intimagen = intimagen + 1
                        FileCopy "\\PCTELEMEMO1\" & Mid(RscomRequer!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\error\" & CStr(intimagen) & ".TIF"
                    Else
                        If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & Directorio_Remito, vbDirectory) = "" Then
                            FileSystem.MkDir "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & Directorio_Remito
                        End If
                        FileCopy "\\PCTELEMEMO1\" & Mid(RscomRequer!Suspense_File, 3), "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & Directorio_Remito & "\" & Imagen
                    End If
                    
                    
                   
                    
                End If
                
                Sql = " Update basasql.dbo.TELEFORM_REQUERIMIENTO"
                Sql = Sql & " SET ENBASA =1"
                Sql = Sql & " Where ID = " & RscomRequer!ID
                ExecutarSql Sql
                
                
    RscomRequer.MoveNext
 Loop

Clipboard.Clear
Clipboard.SetText Remito_Manual_Error
MsgBox "Terminado"

End Sub

Private Sub Command1_Click()

        Dim Sql As String
        Sql = " SELECT ID_CLIENTE_LEGAJO, COD_INDICE, CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE, FECHA_ACTUALIZACION , ID_Personal"
        Sql = Sql + " From LEGAJOS "
        Sql = Sql + " WHERE (COD_CLIENTE = 20) AND (NRO_CAJA IN (26849, 28130,"
        Sql = Sql + vbCrLf & " 28070, 28136, 27959, 26122, 28068, 27841, 27923, 27847,"
        Sql = Sql + vbCrLf & " 28069, 27942, 26167, 27941, 27940, 26165, 27916, 27980,"
        Sql = Sql + vbCrLf & " 28129, 27846, 27848, 28135, 28132, 27944, 27983, 27962,"
        Sql = Sql + vbCrLf & " 27524, 27520, 27969, 27978, 27845, 27917, 27519, 27912,"
        Sql = Sql + vbCrLf & " 28055, 27918, 27956, 28128, 27974, 27840, 28057, 27976,"
        Sql = Sql + vbCrLf & " 28139, 27965, 27603, 27827, 27961, 27911, 27843, 27920,"
        Sql = Sql + vbCrLf & " 27968, 26166, 27838, 28064, 27850, 28141, 28134, 28127,"
        Sql = Sql + vbCrLf & " 27919, 27982, 27970, 28252, 28062, 27921, 27954, 27971,"
        Sql = Sql + vbCrLf & " 28066, 26168, 28056, 28133, 27973, 28073, 27922, 27975,"
        Sql = Sql + vbCrLf & " 28060, 26030, 26170, 26169, 27958, 27849, 26176, 26164,"
        Sql = Sql + vbCrLf & " 27960, 27981, 28256, 27914, 27943, 28144, 27839, 27816,"
        Sql = Sql + vbCrLf & " 28071, 28058, 28143, 28061, 26171, 28142, 28059, 27525,"
        Sql = Sql + vbCrLf & " 27913, 28255, 27842, 27844, 27836, 28065, 27837, 27955,"
        Sql = Sql + vbCrLf & " 27977, 27826, 27616, 26843, 27608, 27825, 28067, 28072,"
        Sql = Sql + vbCrLf & " 27964, 27397, 16730, 28063, 28054, 27979, 27966, 27963,"
        Sql = Sql + vbCrLf & " 27818, 27815, 27800, 27799, 27797, 27796, 27614, 27613,"
        Sql = Sql + vbCrLf & " 27610, 28613, 27607, 27396, 27528, 27604, 27606, 27527,"
        Sql = Sql + vbCrLf & "  16836))"

    Dim rs  As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1


Do While Not rs.EOF
    Sql = " Update OSEP_LEGAJOS_ARCHIVO "
    Sql = Sql + vbCrLf & " SET COD_CLIENTE_LEGAJO = " & rs!ID_CLIENTE_LEGAJO
    Sql = Sql + vbCrLf & " , FECHA_ACTUALIZACION = '20/10/2006'"
    Sql = Sql + vbCrLf & " WHERE LEGAJO = '" & LegajoOrden(rs!CLIENTE_LEGAJO) & "'"
    ExecutarSql Sql
    rs.MoveNext
Loop








End Sub

Private Sub cmdReferenciasCargadasPercistencia_Click()
    Dim Paso As String
        CommonDialog1.FileName = "\\Base\Basa\Usuarios\Clientes\ECOGAS\Actualización de Referencia\Sin Cargar\Remitos\*.XLS"
        CommonDialog1.ShowOpen
        Paso = CommonDialog1.FileName
    If Control_Excel_Referencias_Cargadas(Paso, True) Then
        MsgBox "Existe un errores", vbCritical, "Control Carga"
    Else
        Control_Excel_Referencias_Cargadas Paso, True
    End If
End Sub

Private Sub Command2_Click()
'Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'    Dim B_Error As Boolean
'    Dim Caja As Long
'    Dim Cliente As Integer
'    Dim msgError As String
'    Dim ErrorEstado, Indice As String
'    Dim ErrorRef As String
'    Dim ID_Sql_Imagen As Long
'    Dim ID_SQL As Long
'    ID_Sql_Imagen = 0
'    Dim R As Integer
'    Dim ConAsistir As New ADODB.Connection
'    Set ConAsistir = New ADODB.Connection
'    ConAsistir.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\asistircontrol.mdb;Persist Security Info=False"
'Dim dia As String
'
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Open("c:\HISTORIAS CLINICAS DIARIAS.xls")
'    Set hojaEx = libroEx.Worksheets.Item(1)
'
'
'    Dim C As Integer
'
'
'
'
'
'With hojaEx
'
'   For C = 1 To 300
'        For R = 1 To 10000
'          If Cells(R, C) = "" Then
'            Exit For
'          Else
'            If R = 1 Then
'                Rem fecha
'                dia = Cells(R, C)
'                Else
'                ConAsistir.Execute "insert into control (dia,hc) values ('" & dia & "'," & Datos & ")"
'                MsgBox Cells(R, C)
'            End If
'          End If
'        Next
'    Next
'
'End With
'
'
'
'libroEx.Save
'libroEx.Close
'ApExcel.Quit
'Set hojaEx = Nothing
'Set libroEx = Nothing
'Set ApExcel = Nothing
    
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command9_Click()

End Sub

Private Sub Command3_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

Dim NRO_CAJA, Indice, FECHA_DESDE, FECHA_HASTA As String
Dim LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA As String
Dim Descripcion As String
Dim COD_ID_REFERENCIA As Long
Sql = " SELECT NRO_CAJA, INDICE, FECHA_DESDE, FECHA_HASTA,"
    Sql = Sql & " LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA,"
    Sql = Sql & "  Descripcion"
  Sql = Sql & " From ECOGASREFDUPLICADAS"
  Sql = Sql & " ORDER BY NRO_CAJA"
rs.Open Sql, ConActiva, 0, 1
Exit Sub
COD_ID_REFERENCIA = 409335
Do While Not rs.EOF
    COD_ID_REFERENCIA = COD_ID_REFERENCIA + 1
    Sql = "SELECT *  From REFERENCIAS"
    Sql = " UPDATE REFERENCIAS SET EXPEDIENTE = '03/07/2007'"
    Sql = Sql & " Where COD_CLIENTE = 4 And NRO_CAJA = " & rs!NRO_CAJA & ""
    Sql = Sql & vbCrLf & " AND INDICE ='" & rs!Indice & "'"
    
    
    NRO_CAJA = rs!NRO_CAJA
    Indice = "'" & rs!Indice & "'"
    If IsNull(rs!FECHA_DESDE) Then
        Sql = Sql & vbCrLf & " AND FECHA_DESDE IS NULL "
        FECHA_DESDE = "NULL"
    Else
        FECHA_DESDE = FechaServerTipo(rs!FECHA_DESDE)
        Sql = Sql & vbCrLf & " AND FECHA_DESDE = " & FECHA_DESDE
    End If
    
    If IsNull(rs!FECHA_HASTA) Then
        FECHA_HASTA = "NULL"
        Sql = Sql & vbCrLf & " AND FECHA_HASTA IS NULL "
    Else
        FECHA_HASTA = FechaServerTipo(rs!FECHA_HASTA)
        Sql = Sql & vbCrLf & " AND FECHA_HASTA = " & FECHA_HASTA
    End If
    
    If IsNull(rs!LETRA_DESDE) Then
        LETRA_DESDE = "NULL"
        Sql = Sql & vbCrLf & " AND LETRA_DESDE IS NULL"
    Else
        LETRA_DESDE = "'" & rs!LETRA_DESDE & "'"
        Sql = Sql & vbCrLf & " AND LETRA_DESDE = " & LETRA_DESDE
        
    End If
    
    If IsNull(rs!LETRA_HASTA) Then
        LETRA_HASTA = "NULL"
        Sql = Sql & vbCrLf & " AND LETRA_HASTA IS NULL "
    Else
        LETRA_HASTA = "'" & rs!LETRA_HASTA & "'"
        Sql = Sql & vbCrLf & " AND LETRA_HASTA = " & LETRA_HASTA
    End If
    
    If IsNull(rs!NRO_DESDE) Then
        NRO_DESDE = "NULL"
         Sql = Sql & vbCrLf & " AND NRO_DESDE IS NULL "
        
    Else
        NRO_DESDE = rs!NRO_DESDE
         Sql = Sql & vbCrLf & " AND NRO_DESDE = " & NRO_DESDE
    End If
    
    
    If IsNull(rs!NRO_HASTA) Then
        NRO_HASTA = "NULL"
        Sql = Sql & vbCrLf & " AND NRO_HASTA IS NULL "
    Else
        NRO_HASTA = rs!NRO_HASTA
        Sql = Sql & vbCrLf & " AND NRO_HASTA = " & NRO_HASTA
    End If
    
   If IsNull(rs!Descripcion) Then
        Sql = Sql & vbCrLf & " AND DESCRIPCION IS NULL "
        Descripcion = "NULL"
    Else
        Descripcion = "'" & rs!Descripcion & "'"
        Sql = Sql & vbCrLf & " AND DESCRIPCION =" & Descripcion
    End If
    
    
    Sql = " INSERT INTO REFERENCIAS"
   Sql = Sql & vbCrLf & " (COD_ID_REFERENCIA, COD_CLIENTE, NRO_CAJA,"
   Sql = Sql & vbCrLf & "  COD_TIPO_ALMACENAMIENTO, INDICE, DESCRIPCION,"
    Sql = Sql & vbCrLf & "  FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA,"
    Sql = Sql & vbCrLf & "  LETRA_DESDE, LETRA_HASTA, FECHA_MODIFICACION,"
    Sql = Sql & vbCrLf & "  USUARIO_MODIFICACION, PASOARCHIVO,BORRADO)"
Sql = Sql & vbCrLf & "  Values "
Sql = Sql & vbCrLf & "(" & COD_ID_REFERENCIA & ",4," & NRO_CAJA & ","
Sql = Sql & vbCrLf & "0," & Indice & "," & Descripcion & ","
    Sql = Sql & vbCrLf & FECHA_DESDE & "," & FECHA_HASTA & "," & NRO_DESDE & "," & NRO_HASTA & ","
    Sql = Sql & vbCrLf & LETRA_DESDE & "," & LETRA_HASTA & ","
   Sql = Sql & vbCrLf & "'03/072007','ControlDuplicados', 'ControlDuplicados',0)"

    ExecutarSql Sql
    
    rs.MoveNext
Loop

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    Dim s_ConSql As String
    PasoServer = "\\Serverbackup_1\E\Usuarios\Basa\Operaciones\Referencias\"
    PasoServerImagenSql = "\\Server001\ImagenesSql\"
    s_ConSql = strConBasa
    Set ConSql = New ADODB.Connection
    ConSql.Open s_ConSql
    clienteCajasChicas = 0
End Sub

'Private Sub Actualizar_Documento2()
'    Dim conexel As New ADODB.Connection
'    Dim rsExel As New ADODB.Recordset
'    Dim Caja As Long
'    Dim Cliente As Integer
'    Dim Sql As String
'    Dim ID As Long
'    Dim Indice As String
'    Dim INDICE_ANTERIOR2 As String
'    Dim strCodigo As String
'    Dim rsCodigo As ADODB.Recordset
'    Cliente = ctlClientes.Valor
'Dim R As Integer
'   On Error GoTo salir
'   conexel.Open "Provider=MSDASQL.1;Persist Security Info=False;Mode=ReadWrite;Extended Properties=DBQ=" & CommonDialog1.FileName & ";DefaultDir=F:\Público\1- Osep\Notas y Cartas;Driver={Microsoft Excel Driver (*.xls)};DriverId=790;FIL=excel 5.0;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;ReadOnly=1;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
'   rsExel.Open "SELECT * FROM `Referencias$`", conexel
'    R = 0
'    ConBasa.BeginTrans
'    Do While Not rsExel.EOF And R < 1000
'      R = R + 1
'      If Not IsNull(rsExel!F12) Then
'            If IsNumeric(Replace(rsExel!F12, "'", "")) Then
'                'Indice
'                ID = Replace(rsExel!F12, "'", "")
'                If IsNumeric(rsExel!F1) Then 'Documento
'                     If TraerIndice(rsExel!F1, Cliente) = "" Then
'                        Indice = ""
'                     Else
'                        Indice = "'" & TraerIndice(rsExel!F1, Cliente) & "'"
'                     End If
'                 Else
'                     MsgBox "Error en el documento ID:" & ID & vbCrLf & "El error esta copiado en el Clipboard"
'                     Indice = ""
'                End If
'                INDICE_ANTERIOR2 = Indice_Anterior(ID)
''                If Indice <> "" And INDICE_ANTERIOR2 <> Indice Then
''                        Sql = "   Update REFERENCIAS"
''                        Sql = Sql & vbCrLf & " SET INDICE = " & Indice & ", PASOARCHIVO = '" & CommonDialog1.FileTitle & "', INDICE_ANTERIOR = '" & INDICE_ANTERIOR2 & "',"
''                        Sql = Sql & vbCrLf & " FECHA_MODIFICACION = TO_DATE('" & Format(Now, "DD/MM/YYYY") & "', 'DD/MM/YYYY'),"
''                        Sql = Sql & vbCrLf & " USUARIO_MODIFICACION = 'S_Cambio_indice'"
''                        Sql = Sql & vbCrLf & " Where COD_ID_REFERENCIA = " & ID
''                        Sql = Sql & vbCrLf & " And COD_CLIENTE = " & Cliente
''                        ExecutarSql Sql
''                End If
'
'             End If
'     End If
'       rsExel.MoveNext
'  Loop
'
' ConBasa.CommitTrans
' MsgBox "Proseso Terminado"
' Exit Sub
'salir:
'     MsgBox " Error en la operacion"
'     ConBasa.RollbackTrans
'End Sub
'
Public Function Max_referencia() As Long
    Dim rs As ADODB.Recordset
    Dim Sql As String
        Set rs = New ADODB.Recordset
        Sql = " SELECT    MAX(COD_ID_REFERENCIA) AS MAX_COD_ID_REFERENCIA"
        Sql = Sql & " From REFERENCIAS "
        rs.Open Sql, ConActiva, 0, 1
        Max_referencia = rs!Max_Cod_Id_Referencia + 1
End Function

Public Function InsertarReferencias(COD_CLIENTE As Integer, NRO_CAJA As Long, Indice As String, Descripcion As String, _
FECHA_DESDE As String, FECHA_HASTA As String, NRO_DESDE As String, NRO_HASTA As String, LETRA_DESDE As String, _
LETRA_HASTA As String, EXPEDIENTE As String, PASOARCHIVO As String, ID_imagen As Long, ControlExcel As String, USUARIO_MODIFICACION As String) As Long
Dim Sql As String
Dim FECHA_MODIFICACION As String
Dim COD_ID_REFERENCIA As Long
FECHA_MODIFICACION = SysDateMinutoSegundo
COD_ID_REFERENCIA = Max_referencia


 
    Sql = " INSERT INTO REFERENCIAS"
    Sql = Sql & vbCrLf & " (COD_CLIENTE, NRO_CAJA, INDICE, DESCRIPCION "
    Sql = Sql & vbCrLf & "  , FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA"
    Sql = Sql & vbCrLf & "  , LETRA_DESDE, LETRA_HASTA, FECHA_MODIFICACION "
    Sql = Sql & vbCrLf & "  , USUARIO_MODIFICACION, PASOARCHIVO"
    Sql = Sql & vbCrLf & "  , BORRADO ,ID_IMAGEN,  CONTROLEXCEL,EXPEDIENTE, FK_PERSONAL_CREACION , FK_PERSONAL_MODIFICACION  )"
    Sql = Sql & vbCrLf & " VALUES "
    Sql = Sql & vbCrLf & "(" & COD_CLIENTE & "," & NRO_CAJA & "," & Indice & "," & Descripcion
    Sql = Sql & vbCrLf & " ," & FECHA_DESDE & "," & FECHA_HASTA & "," & Replace(NRO_DESDE, ".", "") & "," & Replace(NRO_HASTA, ".", "")
    Sql = Sql & vbCrLf & "," & LETRA_DESDE & "," & LETRA_HASTA & "," & FECHA_MODIFICACION
    Sql = Sql & vbCrLf & "," & USUARIO_MODIFICACION & "," & PASOARCHIVO
    Sql = Sql & vbCrLf & ",0, " & ID_imagen & "," & ControlExcel & ",'" & EXPEDIENTE & "'," & USUARIO_MODIFICACION & "," & USUARIO_MODIFICACION & " )"
    ExecutarSql Sql
    
    InsertarReferencias = COD_ID_REFERENCIA
  
    
'            Dim conA As New ADODB.Connection
'        conA.Open strConBasa , 0 ,1
'
'        Sql = " Update Documentos "
'        Sql = Sql & " Set Estado = 50 "
'        Sql = Sql & " Where Id = " & ID_imagen
'        conA.Execute Sql
End Function


Public Function NombreArchivo(Paso As String) As String

Dim i As Integer
Dim D As Integer
D = Len(Paso)
For i = 0 To Len(Paso)
    D = D - 1
    Debug.Print Mid(Paso, D + 1)
    If Mid(Paso, D, 1) = "\" Then
         NombreArchivo = Mid(Paso, D + 1)
         Debug.Print Mid(Paso, D + 1)
         Exit Function
    End If
Next

End Function


Public Function ValidarImagenSql(PasoImagen As String) As Long
    Dim s_RsSql As String
    Dim RsSql As ADODB.Recordset
        Set RsSql = New ADODB.Recordset
        s_RsSql = " SELECT ID, Indice "
        s_RsSql = s_RsSql & vbCrLf & " FROM DOCUMENTOS_DIGITALES "
        s_RsSql = s_RsSql & vbCrLf & " WHERE  PasoOrigen = '" & PasoImagen & "'"
        RsSql.Open s_RsSql, ConSql
        If Not RsSql.EOF Then
            ValidarImagenSql = RsSql!ID
        Else
            ValidarImagenSql = 0
        End If
        
End Function


Public Function ControlReferenciaCargada(Cliente As Integer, NRO_CAJA As Long) As String
 Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT COD_CLIENTE, NRO_CAJA "
        Sql = Sql & " From REFERENCIAS  "
        Sql = Sql & " Where COD_CLIENTE = " & Cliente
        Sql = Sql & " And NRO_CAJA = " & NRO_CAJA

    rs.Open Sql, ConActiva, 0, 1
    ControlReferenciaCargada = ""
    If rs.EOF Then
        ControlReferenciaCargada = ""
    Else
        ControlReferenciaCargada = "Referencia ya Cargada"
    End If


End Function

Public Function Control_Excel_Cliente_Documento(Paso As String, ModificarNombre As Boolean) As Boolean

'        Dim ApExcel As Excel.Application
'        Dim libroEx As Excel.Workbook
'        Dim hojaEx As Excel.Worksheet
'        Dim B_Error As Boolean
'        Dim i As Integer
'        Dim msgError  As String
'
'        'abrir hoja excel
'        Set ApExcel = New Excel.Application
'        Set libroEx = Excel.Workbooks.Open(paso)
'        Set hojaEx = libroEx.Worksheets.Item(2)
'
'        'Variables del docuemneto
'        Dim Documento_Actual As String
'        Dim Indice_Actual_Control As String
'        Dim ID_Referencias As String
'        Dim Error_Documento As String
'        Dim COD_CLIENTE As Integer
'        Dim C_ID As Integer
'        Dim C_Documento As Integer
'        Dim C_Error As Integer
'        Dim C_Caja As Integer
'        Dim C_Ok As Integer
'        Dim Contar_Vacias  As Integer
'
'
'        On Error GoTo salir:
'
'        'Control de Nombre de planilla
'
'        Select Case hojaEx.Name
'        Case "Cajas"
'            If Trim(Cells(2, 3)) <> "REFERENCIAS POR CAJAS" Then
'                MsgBox "Error en el formato", vbInformation
'                GoTo salir
'            Else
'                C_ID = 4
'                C_Documento = 2
'                C_Error = 5
'                C_Caja = 1
'                C_Ok = 6
'            End If
'        Case Else
'            MsgBox "El nombre de la planilla no es el correcto", vbInformation
'            GoTo salir
'        End Select
'
'        'Control de Cliente
'        If IsNull(ctlClientes.Valor) Then
'            MsgBox "Ingrese el cliente", vbCritical
'            GoTo salir
'        Else
'            COD_CLIENTE = ctlClientes.Valor
'        End If
'
'    'Control de datos
'
'    With hojaEx
'            For i = 4 To 1500
'               Indice_Actual_Control = ""
'               Error_Documento = ""
'               ID_Referencias = 0
'               If Len(Cells(i, C_Caja)) = 0 Then
'                    If Not IsNumeric(Cells(i, C_Caja)) Or Cells(i, C_Caja) = "" Then
'                            If Cells(i, C_Documento) > 1 Then
'                                    Contar_Vacias = 0
'                                    Indice_Actual_Control = TraerIndice(CInt(Cells(i, C_Documento)), COD_CLIENTE)
'                                    If Indice_Actual_Control = "" Then
'                                        Error_Documento = Indice_Actual_Control
'                                    End If
'                                    ID_Referencias = Control_ID_Referencias(COD_CLIENTE, CLng(Replace(Cells(i, C_ID), "ID:", "")))
'                                    If ID_Referencias = "" Then
'                                        Error_Documento = "El ID Referencia No corresponde"
'                                    End If
'                            End If
'                    End If
'                Else
'                    Contar_Vacias = Contar_Vacias + 1
'                    If Contar_Vacias > 10 Then
'                        Exit For
'                    End If
'                End If
'
'                'Control de Error
'                If Error_Documento <> "" Then
'                    Cells(i, C_Error) = Error_Documento
'                    B_Error = True
'                Else
'                    If ID_Referencias <> 0 Then
'                        Cells(i, C_Ok) = "OK"
'                    Else
'                        Cells(i, C_Ok) = "Sin Asisgnar"
'                    End If
'                End If
'                lblTitulo.Caption = "Control registros"
'                lblCantidadRegistros.Caption = i
'                lblCantidadRegistros.Refresh
'
'            Next
'End With
'
'        libroEx.Save
'        libroEx.Close
'        ApExcel.Quit
'        Set hojaEx = Nothing
'        Set libroEx = Nothing
'        Set ApExcel = Nothing
'        Control_Excel_Cliente_Documento = B_Error
'        If ModificarNombre Then
'            If B_Error = False Then
'                FileSystem.FileCopy paso, Mid(paso, 1, Len(paso) - 4) & " Control_OK " & " .xls"
'                FileSystem.Kill paso
'            End If
'        End If
'Exit Function
'salir:
'libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            Control_Excel_Cliente_Documento = True

End Function

Public Function Control_ID_Referencias(COD_CLIENTE As Integer, COD_ID_REFERENCIA As Long) As String
    Dim rs As ADODB.Recordset
    Dim Sql As String
        Set rs = New ADODB.Recordset
            Sql = " SELECT COD_CLIENTE, COD_ID_REFERENCIA"
            Sql = Sql & " From REFERENCIAS"
            Sql = Sql & "  Where Cod_Cliente = " & COD_CLIENTE
            Sql = Sql & "  And COD_ID_REFERENCIA = " & COD_ID_REFERENCIA
            rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            Control_ID_Referencias = ""
        Else
            Control_ID_Referencias = rs!COD_ID_REFERENCIA
        End If


End Function
Public Function Indice_Anterior(COD_CLIENTE As Integer, COD_ID_REFERENCIA As Long) As String
    Dim rs As ADODB.Recordset
    Dim Sql As String
        Set rs = New ADODB.Recordset
            Sql = " SELECT COD_CLIENTE, COD_ID_REFERENCIA,INDICE"
            Sql = Sql & " From REFERENCIAS"
            Sql = Sql & "  Where Cod_Cliente = " & COD_CLIENTE
            Sql = Sql & "  And COD_ID_REFERENCIA = " & COD_ID_REFERENCIA
            rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            Indice_Anterior = ""
        Else
            Indice_Anterior = rs!Indice
        End If


End Function


Public Sub PercistenciaLegajosOsde(PasoOrigen As String)
 Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim i As Integer
        Dim NRO_CAJA As String
        Dim Legajo As String
        Dim Sql As String
        'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(PasoOrigen)
        Set hojaEx = libroEx.Worksheets.Item(1)
        Dim rs As ADODB.Recordset
        
        For i = 2 To 6000
               'Control de fin de rows
            If hojaEx.Cells(i, 1) = "" And hojaEx.Cells(i, 2) = "" Then
                Exit For
            Else
            
            Set rs = New ADODB.Recordset
                Sql = " SELECT * "
                Sql = Sql + " From OSEP_LEGAJOS_ARCHIVO "
                Sql = Sql + " WHERE LEGAJO ='" & Trim(hojaEx.Cells(i, 1)) & "'"
            rs.Open Sql, ConActiva, 0, 1
            If rs.EOF Then
                    
                                      
                    
                    Sql = " INSERT INTO OSEP_LEGAJOS_ARCHIVO "
                    Sql = Sql + " (NRO_CAJA, LEGAJO)"
                    If hojaEx.Cells(i, 2) <> "" Then
                        NRO_CAJA = hojaEx.Cells(i, 2)
                    Else
                        NRO_CAJA = "NULL"
                    End If
                    Sql = Sql + " VALUES (" & NRO_CAJA & ",'" & Trim(hojaEx.Cells(i, 1)) & "')"
                    ExecutarSql Sql
            Else
                If hojaEx.Cells(i, 2) <> "" Then
                    Sql = " UPDATE OSEP_LEGAJOS_ARCHIVO "
                    Sql = Sql + " Set NRO_CAJA = " & hojaEx.Cells(i, 2)
                    Sql = Sql + " WHERE LEGAJO ='" & Trim(hojaEx.Cells(i, 1)) & "'"
                    ExecutarSql Sql
                End If
            End If
           End If
        lblCantidadRegistros.Caption = i
        lblCantidadRegistros.Refresh
        
       Next
         MsgBox "La grabación de realizo con exito", vbInformation
        libroEx.Save
        libroEx.Close
        ApExcel.Quit
        Set hojaEx = Nothing
        Set libroEx = Nothing
        Set ApExcel = Nothing
        FileSystem.FileCopy PasoOrigen, Mid(PasoOrigen, 1, Len(PasoOrigen) - 4) & " Procesado" & " .xls"
        FileSystem.Kill PasoOrigen
End Sub

Public Function LegajoOrden(DATO As String) As String
Dim i As Integer
Dim DATO2 As String

For i = 1 To Len(DATO)
If Mid(DATO, i, 1) <> "-" Then
    If Mid(DATO, i, 1) = "0" Then
    
    Else
       DATO2 = Mid(DATO, i)
       Exit For
    End If
 Else
    Exit For
 End If
 
Next
    LegajoOrden = Replace(DATO2, " ", "")
    
End Function

Public Function ControlCaja(Caja As Long, Cliente As Integer, Indice As String, fechadesde As String, NRODESDE As String, Descripcion As String) As String
    
    ControlCaja = ""
    Dim rsCaja As New ADODB.Recordset
    Dim Sql As String
    On Error GoTo salir:
    Sql = " SELECT     NRO_CAJA, FK_CLIENTE"
    Sql = Sql & vbCrLf & " From dbo.Cajas "
    Sql = Sql & vbCrLf & " Where FK_CLIENTE = " & Cliente
    Sql = Sql & vbCrLf & " And NRO_CAJA = " & Caja
    rsCaja.Open Sql, ConActiva, 0, 1
    
    If rsCaja.EOF Then
        ControlCaja = "La caja No esta asignada al cliente"
        Exit Function
    End If
    
    
    Set rsCaja = New ADODB.Recordset
    Sql = " SELECT  * From dbo.REFERENCIAS"
    Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & vbCrLf & " And NRO_CAJA = " & Caja
    
    If Indice <> "Error" Then
    If Indice <> 0 Then
        Sql = Sql & vbCrLf & " And INDICE = '" & Indice & "'"
    End If
    End If
    If fechadesde <> "0" Then
        Sql = Sql & vbCrLf & " And FECHA_DESDE = " & FechaFormato(fechadesde)
    End If
        
    If NRODESDE <> "0" Then
        If IsNumeric(NRODESDE) Then
        Sql = Sql & vbCrLf & " And NRO_DESDE = '" & NRODESDE & "'"
        End If
    End If

 If Descripcion <> "" Then
        
        Sql = Sql & vbCrLf & " And descripcion like  '%" & Trim(Descripcion) & "%'"
        
    End If
    
    
    rsCaja.Open Sql, ConActiva, 0, 1
    
    If chkNoControlRefCargada.value = 0 Then
        If Not rsCaja.EOF Then
            ControlCaja = "La caja tiene referencia identica"
            Exit Function
            End If
    End If
    Exit Function
salir:
    ControlCaja = "Error verifique"
End Function

Public Function ControlReferencia(Caja As Long, Cliente As Integer) As String
    
    ControlReferencia = ""
    Dim rsCaja As New ADODB.Recordset
    Dim Sql As String
    On Error GoTo salir:
    Sql = " SELECT     NRO_CAJA, FK_CLIENTE"
    Sql = Sql & vbCrLf & " From dbo.Cajas "
    Sql = Sql & vbCrLf & " Where FK_CLIENTE = " & Cliente
    Sql = Sql & vbCrLf & " And NRO_CAJA = " & Caja
    rsCaja.Open Sql, ConActiva, 0, 1
    
    If rsCaja.EOF Then
        ControlReferencia = "La caja No esta asignada al cliente"
        Exit Function
    End If
    
    
    Set rsCaja = New ADODB.Recordset
    Sql = " SELECT  * From dbo.REFERENCIAS"
    Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & vbCrLf & " And NRO_CAJA = " & Caja
     rsCaja.Open Sql, ConActiva, 0, 1
    If Not rsCaja.EOF Then
        ControlReferencia = "La caja tiene referencia "
    End If
    Exit Function
salir:
    ControlReferencia = "Errror en caja cliente "
End Function

Public Function VerificarDocumentoLegajo(Cliente As Integer, Documento As Integer, Indice As String, FK_Indice As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    VerificarDocumentoLegajo = False
    Sql = " SELECT     COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, ID, TIPO_INDICE"
Sql = Sql & " From INDICES"
Sql = Sql & "  Where COD_CLIENTE =  " & Cliente
Sql = Sql & "  And ID_CODIGO_DOCUMENTO = " & Documento
    
    rs.Open Sql, ConActiva, 0, 1
    
    If Not rs.EOF Then
    
        If Trim(rs!Tipo_Indice) = "Legajo" Then
             VerificarDocumentoLegajo = True
              Indice = "'" & Trim(rs!Indice) & "'"
              FK_Indice = rs!ID
       End If
    
    End If
    
     Debug.Assert VerificarDocumentoLegajo = True


End Function

Public Function controlEtiqueta(Etiqueta As Long) As String
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " SELECT     ID_LEGAJO, NRO_DESDE ,  NRO_CAJA"
    Sql = Sql & " From LEGAJOS "
    Sql = Sql & "  Where ID_LEGAJO = " & Etiqueta
    Sql = Sql & " AND NOT (NRO_CAJA IS NULL) "
    rs.Open Sql, ConActiva, 0, 1

        controlEtiqueta = ""
If Not rs.EOF Then

                controlEtiqueta = " Etiqueta en uso "

End If






End Function

Public Function Percistencia_legajos_supervielle(Paso As String, ModificarNombre As Boolean) As Boolean
'Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'
'    Percistencia_legajos_supervielle = True
'    Dim C_Error As Integer
'    Dim C_Documento As Integer
'    Dim C_Caja As Integer
'    Dim C_Etiqueta As Integer
'    Dim C_Tipo_Documento As Integer
'    Dim C_Nro_Documento As Integer
'    Dim C_Apellido_Nombre As Integer
'    Dim C_Descripcion As Integer
'    Dim C_Fecha_Carga   As Integer
'    Dim C_Personal_Carga As Integer
'    Dim C_Fecha_desde As Integer
'    Dim C_Fecha_hasta As Integer
'
'
'       C_Error = 1
'    C_Documento = 2
'    C_Caja = 3
'    C_Etiqueta = 4
'    C_Tipo_Documento = 5
'    C_Nro_Documento = 6
'    C_Apellido_Nombre = 7
'    C_Descripcion = 8
'    C_Fecha_Carga = 9
'    C_Personal_Carga = 10
'    C_Fecha_desde = 11
'    C_Fecha_hasta = 12
'
'
'
'    Dim FK_INDICES  As Long
'    Dim Cod_Indice As String
'    Dim NRO_DESDE As String
'    Dim NRO_HASTA As String
'    Dim LETRA_DESDE As String
'    Dim LETRA_HASTA As String
'    Dim FECHA_DESDE As String
'    Dim FECHA_HASTA As String
'    Dim FECHA_CREACION   As String
'    Dim NRO_CAJA As String
'    Dim COD_CLIENTE  As String
'    Dim Descripcion  As String
'    Dim Cod_Estado  As String
'    Dim ARCHIVO_POR_LOTE As String
'    Dim sql As String
'    Dim i As Long
'    Dim R As Integer
'
'
'
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'     Set libroEx = Excel.Workbooks.Open(paso)
'    Set hojaEx = libroEx.Worksheets.Item(1)
'
'
'    Dim C As Integer
'
'
'
'
'    If UCase(hojaEx.Name) <> "CARGA_LEGAJOS" Then
'        MsgBox "El nombre de la planilla no es el correcto", vbInformation
'        libroEx.Close
'        ApExcel.Quit
'        Set hojaEx = Nothing
'        Set libroEx = Nothing
'        Set ApExcel = Nothing
'        Exit Function
'    End If
'
'
'With hojaEx
'
'    COD_CLIENTE = ctlClientes.Valor
'    Cod_Estado = 2
'    ARCHIVO_POR_LOTE = "'" & Trim(NombreArchivo(paso)) & "'"
'
'
'
'    For i = 2 To InputBox("Ingrese la cantidad de registros") + 50
'
'    If .Cells(i, C_Error) = "" Then
'        Rem .Cells(i, C_Error) = ""
'      'Control de fin de rows
'       If .Cells(i, C_Documento) = "" And .Cells(i, C_Caja) = "" Then
'            .Cells(i, C_Error) = "No se registro"
'
'       Else
'
'
'        FK_INDICES = 0
'        Cod_Indice = ""
'
'
'        If .Cells(i, C_Nro_Documento) = "NO TIENE" Or Trim(.Cells(i, C_Nro_Documento)) = "" Then
'            NRO_DESDE = 0
'            NRO_HASTA = 0
'        Else
'            NRO_DESDE = .Cells(i, C_Nro_Documento)
'            NRO_HASTA = .Cells(i, C_Nro_Documento)
'        End If
'
''        LETRA_DESDE = "'" & Replace(Trim(.Cells(i, C_Apellido_Nombre)), "'", "´") & "'"
''        LETRA_HASTA = "'" & Trim(.Cells(i, C_Tipo_Documento)) & "'"
''
'        LETRA_DESDE = "'" & UCase(Replace(Trim(.Cells(i, C_Apellido_Nombre)), "'", "´")) & "'"
'        LETRA_HASTA = "'" & UCase(Replace(Trim(.Cells(i, C_Tipo_Documento)), "'", "´")) & "'"
'
'        NRO_CAJA = .Cells(i, C_Caja)
'
'        If IsDate(.Cells(i, C_Fecha_desde)) Then
'            FECHA_DESDE = "'" & .Cells(i, C_Fecha_desde) & "'"
'        Else
'            FECHA_DESDE = "Null"
'        End If
'
'
'        If IsDate(.Cells(i, C_Fecha_hasta)) Then
'
'            FECHA_HASTA = "'" & .Cells(i, C_Fecha_hasta) & "'"
'        Else
'            FECHA_HASTA = "Null"
'        End If
'
'           ' DECRIPCION
'
'            If Trim(.Cells(i, C_Descripcion)) <> "" Then
'             Descripcion = "'" & UCase(Trim(.Cells(i, C_Descripcion))) & "'"
'            Else
'
'            Descripcion = "Null"
'            End If
'
'
'            Rem Documento
'            If IsNumeric(.Cells(i, C_Documento)) Then
'                  If VerificarDocumentoLegajo(CInt(COD_CLIENTE), .Cells(i, C_Documento), Cod_Indice, FK_INDICES) = False Then
'                   .Cells(i, C_Error) = "El Nro documento no es un Legajo"
'
'                  End If
'
'            Else
'               .Cells(i, C_Error) = "El Nro documento no es un numero"
'
'            End If
'
'
'         FECHA_CREACION = "'" & .Cells(i, C_Fecha_Carga) & "'"
'
'
'        If Trim(NRO_DESDE) = "NO TIENE" Or NRO_DESDE = "NO  TIENE" Then
'        NRO_DESDE = 0
'        NRO_HASTA = 0
'        End If
'
'
'
'        sql = " Update LEGAJOS"
'        sql = sql & " SET FK_INDICES =" & FK_INDICES
'        sql = sql & " , COD_INDICE =" & Cod_Indice
'        sql = sql & " , NRO_DESDE =" & NRO_DESDE
'        sql = sql & " , NRO_HASTA =" & NRO_HASTA
'        sql = sql & " , LETRA_DESDE =" & LETRA_DESDE
'        sql = sql & " , LETRA_HASTA =" & LETRA_HASTA
'        sql = sql & " , FECHA_DESDE =" & FECHA_DESDE
'        sql = sql & " , FECHA_HASTA =" & FECHA_HASTA
'        sql = sql & " , FECHA_CREACION  =" & FECHA_CREACION
'        sql = sql & " , FK_PERSONAL_CREACION =99 " & .Cells(i, C_Personal_Carga)
'        sql = sql & " , NRO_CAJA =" & NRO_CAJA
'        sql = sql & " , COD_CLIENTE =" & COD_CLIENTE
'        sql = sql & " , DESCRIPCION =" & Descripcion
'        sql = sql & " , COD_ESTADO =2"
'        sql = sql & " , ARCHIVO_POR_LOTE= " & ARCHIVO_POR_LOTE
'        sql = sql & " Where ID_LEGAJO =" & .Cells(i, C_Etiqueta)
'        sql = sql & " AND NRO_CAJA IS NULL"
'         R = ExecutarSql(sql)
'
'If R = 0 Then
'    MsgBox "eRROR"
'    .Cells(i, C_Error) = " no Registrado"
'    Debug.Print .Cells(i, C_Etiqueta)
'Else
'.Cells(i, C_Error) = "Registrado"
'End If
'
'
'        Rem Legajos_RecalcularCaracteres_DescripcionRemito (" Where ID_LEGAJO =" & .Cells(i, C_Etiqueta))
'        Debug.Print i
'
'        End If
'
'          .Cells(i, C_Error) = "Registrado"
'        Else
'
'        End If
'
'    Next
'
'End With
'
'libroEx.Save
'libroEx.Close
'ApExcel.Quit
'Set hojaEx = Nothing
'Set libroEx = Nothing
'Set ApExcel = Nothing
'    If ModificarNombre Then
'        If Percistencia_legajos_supervielle = True Then
'            FileSystem.FileCopy paso, Mid(paso, 1, Len(paso) - 4) & " Control_OK " & " .xls"
'            FileSystem.Kill paso
'        End If
'    End If
End Function

Public Function ControlPersonal(ID_Personal As Integer) As Boolean

    ControlPersonal = True
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
    
    
    Sql = " SELECT     IDPERSONAL, NOMBRE, APELLIDO"
Sql = Sql & " FROM Personal"
Sql = Sql & " Where IDPERSONAL = " & ID_Personal


rs.Open Sql, ConActiva, 0, 1

    If rs.EOF Then
        ControlPersonal = False
    
        
    Else
    
    ControlPersonal = True
    
    End If



End Function

Public Function EsCajaCliente(Cliente As Integer, Caja As Long) As String
    Dim rs As New ADODB.Recordset
    
    rs.Open " SELECT * From dbo.Cajas Where FK_CLIENTE =" & Cliente & " And NRO_CAJA = " & Caja, ConActiva, 0, 1
    
    If rs.EOF Then
        EsCajaCliente = "El Cliente no"
    Else
        EsCajaCliente = ""
    End If
    
End Function

Public Function ControlDocumento(Cliente As Integer, Numero_Doc As String) As String
 Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        Dim Sql As String
        
        If Not IsNumeric(Numero_Doc) Then
             ControlDocumento = "En N° de documento no existe"
             Exit Function
        End If
        
        
        Sql = " SELECT *  From INDICES  WHERE COD_CLIENTE =" & Cliente & " AND  ID_CODIGO_DOCUMENTO = " & Numero_Doc
        rs.Open Sql, ConActiva, 0, 1
        
        If rs.EOF Then
            ControlDocumento = "No existe el Documento"
        Else
             ControlDocumento = ""
        End If
End Function


Public Function Control_Documento_campos(Cliente As Integer, Numero_Doc As String, Campo As String) As String
 Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        Dim Sql As String
        
        If Not (IsNumeric(Numero_Doc)) Then
            Control_Documento_campos = "N° de incide "
            Exit Function
        End If
        
        
        
        Sql = " SELECT        COD_CLIENTE, ID_CODIGO_DOCUMENTO, REQUERIR_FECHA_DESDE ,  REQUERIR_LETRA_DESDE, REQUERIR_NRO_DESDE ,  REQUERIR_DESCRIPCION"
        Sql = Sql & "  From INDICES "
        Sql = Sql & " Where COD_CLIENTE = " & Cliente
        Sql = Sql & " And ID_CODIGO_DOCUMENTO =  " & Numero_Doc
        rs.Open Sql, ConActiva, 0, 1
        Control_Documento_campos = ""
        If rs.EOF Then
            Control_Documento_campos = "No existe el Documento"
        Else
            Select Case UCase(Campo)
            Case "FECHA"
                If rs!REQUERIR_FECHA_DESDE = True Then
                    Control_Documento_campos = "La fecha es requerida"
                End If
            Case "NUMERO"
                If rs!REQUERIR_NRO_DESDE = True Then
                    Control_Documento_campos = "El Numero es requerido"
                End If
            
            Case "LETRA"
                If rs!REQUERIR_LETRA_DESDE = True Then
                    Control_Documento_campos = "La Letra es requerida"
                End If
            
            Case "DESCRIPCION"
            
            
            If rs!REQUERIR_DESCRIPCION = 1 Then
                    Control_Documento_campos = "La descripcion es requerida"
                End If
           
            
            End Select
        End If
End Function

Public Function ControlFechas(fechadesde As String, FechaHasta As String) As String


If fechadesde = "" Then
    If FechaHasta <> "" Then
       ControlFechas = "Error  en Fecha"
       Exit Function
    End If
Else
    If IsDate(fechadesde) And IsDate(FechaHasta) Then
        If CDate(fechadesde) > CDate(FechaHasta) Then
            ControlFechas = "La Fecha desde es mayor que la fecha hasta"
            Exit Function
        End If
         If DateDiff("d", Now, FechaHasta) >= 1 Then
            ControlFechas = "La Fecha es futura "
            Exit Function
         End If
    Else
       ControlFechas = "Error  en Fecha"
       Exit Function
    End If
End If

        
End Function

Public Function ControlNumeroLetra(Desde As String, Hasta As String) As String
' Numero
    If Desde = "" And Hasta = "" Then
        Exit Function
    End If
    
    If IsNumeric(Desde) Then
        If Not IsNumeric(Hasta) Then
             ControlNumeroLetra = "Error en Numero Hasta"
             Exit Function
        End If
 
        If CDbl(Desde) > CDbl(Hasta) Then
            ControlNumeroLetra = "Error en Numero desde en mayor  hasta"
            Exit Function
        End If
       End If

'Letra
If Not IsNumeric(Desde) Then
    If Hasta = "" Then
        ControlNumeroLetra = "Error en Letra Hasta"
        Exit Function
    End If
    
End If

End Function

Public Function ControlDesdeHasta(Desde As String, Hasta As String) As String
    If Desde = "" And Desde = "" Then
        Exit Function
    End If
    
    If IsNumeric(Desde) Then
        If Not IsNumeric(Hasta) Then
           ControlDesdeHasta = "Ingrese el Numero Hasta"
        
        End If
        
    End If
    

End Function

Public Function BorrarCajasPlanillas(Paso As String)

'        Dim ApExcel As Excel.Application
'        Dim libroEx As Excel.Workbook
'        Dim hojaEx As Excel.Worksheet
'        Dim Caja As String
'        Dim CajaAnt As Long
'
'       On Error GoTo salir
'       'abrir hoja excel
'        Set ApExcel = New Excel.Application
'        Set libroEx = Excel.Workbooks.Open(paso)
'        Set hojaEx = libroEx.Worksheets.Item(1)
'
'        CajaAnt = 0
'        'Control de Nombre de planilla
'        If hojaEx.Name <> "Planilla para envio por correo" Then
'            MsgBox "El nombre de la planilla no es el correcto" & vbCrLf & " El Nombre correcto es:" & "Planilla para envio por correo", vbInformation
'            Rem Control_Excel_Cliente = True
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            Exit Function
'        End If
'
'
'        If IsNull(ctlClientes.Valor) Then
'            MsgBox "Ingrese el cliente", vbCritical
'           Rem  Control_Excel_Cliente = True
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            Exit Function
'        End If
'
'        If Cells(6, 2) <> "Caja" Then
'            MsgBox "Formato de planilla incorrecto", vbCritical
'            libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            Exit Function
'        End If
'
'        hojaEx.Unprotect 21877471
'        'Iniciaclizacion de Bandera
'
'        Dim FinPlanilla As Boolean
'        Dim ContarBlanco As Integer
'        FinPlanilla = False
'        ContarBlanco = 0
'        Dim i As Integer
'
'        With hojaEx
'        For i = 7 To 6000
'               'Control de fin de rows
'                If Cells(i, 2) = "" Or Cells(i, 3) = "" Then
'                    FinPlanilla = True
'                    ContarBlanco = ContarBlanco + 1
'                    If ContarBlanco > 20 Then
'                        Exit For
'                    End If
'                 Else
'                 If CajaAnt <> CLng(Cells(i, 2)) Then
'                     Caja = Caja & "," & CStr(Cells(i, 2))
'                     CajaAnt = CLng(Cells(i, 2))
'                Else
'
'
'                End If
'
'                End If
'            Next
'
'        End With
'        hojaEx.Protect 21877471
'        Dim regi As Integer
'        Dim sql As String
'        If MsgBox("Usted quiere borrar las cajas " & vbCrLf & Mid(Caja, 2) & vbCrLf & " del cliente " & vbCrLf & ctlClientes.Descripcion, vbYesNo) = vbYes Then
'            sql = " DELETE FROM REFERENCIAS "
'            sql = sql & " WHERE  COD_CLIENTE = " & ctlClientes.Valor
'            sql = sql & "  AND (NRO_CAJA IN (" & Mid(Caja, 2) & " ))"
'             regi = ExecutarSql(sql)
'            MsgBox "Se Borraron " & regi
'
'        End If
'
'
'        libroEx.Saved = True
'        libroEx.Close
'        ApExcel.Quit
'        Set hojaEx = Nothing
'        Set libroEx = Nothing
'        Set ApExcel = Nothing
'
'
'        Exit Function
'salir:
'           hojaEx.Protect 21877471
'           MsgBox "Error en la planilla"
'         Rem   Control_Excel_Cliente = True
'           libroEx.Close
'            ApExcel.Quit
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'End


End Function

Public Sub Control_Referencias(Paso As String)

'Variables de la Planilla
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet
    
    Dim C_Error As Integer
    Dim C_Caja As Integer
    Dim C_Indice As Integer
    Dim C_Etiqueta As Integer
    Dim C_Fecha_desde As Integer
    Dim C_Fecha_hasta As Integer
    Dim C_N°_Desde As Integer
    Dim C_N°_Hasta As Integer
    Dim C_Letra_Desde As Integer
    Dim C_Letra_Hasta As Integer
    Dim C_Descripcion As Integer
    Dim C_Lote As Integer
    Dim C_Correo As Integer
    Dim C_Cliente As Integer
    Dim C_Doc_Cliente As Integer
    Dim C_TipoReferencia As Integer
    

    
    Dim ErrorGeneral As Boolean
    Dim strError As String
    Dim i As Long
    Dim C As Integer
    Dim ControlFin As Integer
    Dim VcontrolEtiqueta As String
    
    ErrorGeneral = False

    Dim FK_INDICES  As Long
    Dim Cod_Indice As String
    Dim NRO_DESDE As String
    Dim NRO_HASTA As String
    Dim LETRA_DESDE As String
    Dim LETRA_HASTA As String
    Dim FECHA_DESDE As String
    Dim FECHA_HASTA As String
    Dim FECHA_CREACION   As String
    Dim NRO_CAJA As String
    Dim COD_CLIENTE  As String
    Dim Descripcion  As String
    Dim Cod_Estado  As String
    Dim ARCHIVO_POR_LOTE As String
    Dim ID_referencia As Long
    Dim Sql As String
    Dim R As Long
    Dim PasoError  As String

PasoError = "Z:\Referencias\CON ERROR\"

    C_Error = 1
    C_Caja = 2
    C_Etiqueta = 3
    C_Indice = 4
    C_Descripcion = 5
    C_Fecha_desde = 6
    C_Fecha_hasta = 7
    C_N°_Desde = 8
    C_N°_Hasta = 9
    C_Letra_Desde = 10
    C_Letra_Hasta = 11
    C_Lote = 12
    
    C_Correo = 13
    C_Cliente = 14
    C_Doc_Cliente = 15
    C_TipoReferencia = 16
    
    Dim FK_CLIENTE As Long
    
    Dim Control_fecha As String
    Dim Control_Numero As String
    Dim Control_Indice As String
    Dim Control_Descripcion As String
    Dim COntroDIfere  As String


On Error GoTo salir
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Open(Paso)
    Set hojaEx = libroEx.Worksheets.Item(1)
    
   hojaEx.Columns(1).Select
   
   libroEx.Unprotect 21877471
   hojaEx.Unprotect 21877471
   
   libroEx.Unprotect 2338
   hojaEx.Unprotect 2338
   
   
clienteCajasChicas = 0

 With hojaEx
        If Not (.Cells(1, 1) = "PLR-002" And UCase(.Cells(6, 2)) = "CAJA" And UCase(Trim(.Cells(6, 11))) = "LETRA HASTA") Then
            
           
            
             libroEx.Close
            ApExcel.Quit
            Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
            MsgBox "La planilla no es Corecta", vbInformation
            Exit Sub
        End If
        
        
        
   .Cells(6, 13) = "Correo"
   .Cells(6, 14) = "Cliente"
   .Cells(6, 15) = "Doc_Cliente"
   .Cells(6, 16) = "Tipo Referencia"
   .Cells(6, 17) = "Lote teleform / Imagen"
   For i = 7 To 30000
             
  

                    strError = ""
                    ' Control estado cajas
principio:
                   If .Cells(i, C_Error) = "REFE" Or .Cells(i, C_Error) = "LEGA" Then
                    GoTo Controlado
                   End If
                    
              
                    
                    If Trim(.Cells(i, C_Caja)) = "" And Trim(.Cells(i, C_Indice)) = "" Then
                    Debug.Print ControlFin
                        If ControlFin > 50 Then
                            Exit For
                           Else
                                i = i + 1
                                ControlFin = ControlFin + 1
                          GoTo principio
                        End If
                    End If


                    If Not IsNumeric(.Cells(i, C_Caja)) Then
                        strError = "La Caja No es un N°"
                    Else
                    
                    
                   
                    
                    
                    FK_CLIENTE = BuscarCliente(.Cells(i, C_Caja), Cells(i, C_Cliente))
                     If FK_CLIENTE = 0 Then
                        strError = "La caja no esta asignada"
                        GoTo Proximo
                     End If
                     
                      Cells(i, C_Cliente) = FK_CLIENTE
                     .Cells(i, C_Correo) = ControlCorreo(CInt(FK_CLIENTE), .Cells(i, C_Caja))
                     .Cells(i, C_Doc_Cliente) = Doc_Usuario(CInt(FK_CLIENTE), .Cells(i, C_Caja))
                        
                        If chkNoControlasEstado.value = 0 Then
                            If Control_Estado_Caja(CInt(FK_CLIENTE), .Cells(i, C_Caja)) <> "" Then
                                strError = Control_Estado_Caja(CInt(FK_CLIENTE), .Cells(i, C_Caja))
                            End If
                        End If
                            
                         If chkBorrarReferencia.value = 1 Then
                            BorrarReferenciaAdministradaPorCLiente CInt(FK_CLIENTE), .Cells(i, C_Caja)
                         End If
                            
                        If chkCambio_Referencia.value = 1 Then
                            BorrarReferencias CInt(FK_CLIENTE), .Cells(i, C_Caja)
                        End If
                            
                        If Trim(.Cells(i, C_TipoReferencia)) = "" Then
'
                    Rem no tocar mada
                        FECHA_DESDE = Trim(.Cells(i, C_Fecha_desde))
                        FECHA_DESDE = Replace(FECHA_DESDE, "-", "/")
                        .Cells(i, C_Fecha_desde).NumberFormat = "@"
                        .Cells(i, C_Fecha_desde) = Trim(CStr(FECHA_DESDE))
                        
                        FECHA_HASTA = Trim(.Cells(i, C_Fecha_hasta))
                        FECHA_HASTA = Replace(FECHA_HASTA, "-", "/")
                        .Cells(i, C_Fecha_hasta).NumberFormat = "@"
                        .Cells(i, C_Fecha_hasta) = Trim(CStr(FECHA_HASTA))
                        Rem hasta aca
                            
                                               
                                            
                                                If .Cells(i, C_N°_Desde) <> "" Then
                                                    Control_Numero = .Cells(i, C_N°_Desde)
                                                Else
                                                    Control_Numero = 0
                                                End If
                                                
                                                If IsNumeric(.Cells(i, C_Indice)) Then
                                                  Control_Indice = BuscarIndice(CInt(FK_CLIENTE), .Cells(i, C_Indice))
                                                Else
                                                   Control_Indice = 0
                                                End If
                                                
                                                If Trim(.Cells(i, C_Descripcion)) <> "" Then
                                                    Control_Descripcion = Trim(.Cells(i, C_Descripcion))
                                                Else
                                                    Control_Descripcion = ""
                                                End If
                                                
                                                
                                            If chkNoControlRefCargada.value = 1 Then
                                             strError = strError & ControlReferencia(.Cells(i, C_Caja), CInt(FK_CLIENTE))
                                            End If
                                            
                                            
                                            
                                            strError = strError & ControlCaja(.Cells(i, C_Caja), CInt(FK_CLIENTE), Control_Indice, FECHA_DESDE, Control_Numero, Control_Descripcion)
                                      
                                        
                                        
                                        
                                        ' Control Etiquetas
                                      
                                      If .Cells(i, C_Etiqueta) <> "" Then
                                            If Not IsNumeric(.Cells(i, C_Etiqueta)) Then
                                                strError = strError & "  " & " La etiquera no es N°"
                                            Else
                                                If controlEtiqueta(.Cells(i, C_Etiqueta)) <> "" Then
                                                    strError = strError & "  " & controlEtiqueta(.Cells(i, C_Etiqueta))
                                                End If
                                            End If
                                        End If
                                        ' Control Indice
                                        If Not IsNumeric(.Cells(i, C_Indice)) Then
                                            strError = strError & "  " & "El Indice no es N°"
                                        Else
                                            If ControlDocumento(CInt(FK_CLIENTE), .Cells(i, C_Indice)) <> "" Then
                                                strError = strError & "  " & ControlDocumento(CInt(FK_CLIENTE), .Cells(i, C_Indice))
                                            End If
                                        End If
                                       ' Control Fecha
                                       
                                      
                                       
                                       
                                      
                                       If Len(FECHA_DESDE) = 4 Then
                                        FECHA_DESDE = "01/01/" & FECHA_DESDE
                                        .Cells(i, C_Fecha_desde) = CStr(FECHA_DESDE)
                                       End If
                                       
                                       If Len(FECHA_HASTA) = 4 Then
                                        FECHA_HASTA = "31/12/" & FECHA_HASTA
                                        .Cells(i, C_Fecha_hasta) = CStr(FECHA_HASTA)
                                       End If
                                       
                                       
                                        If Trim(.Cells(i, C_Fecha_hasta)) = "" Then
                                            If Trim(FECHA_DESDE) <> "" Then
                                                .Cells(i, C_Fecha_hasta) = FECHA_DESDE
                                            End If
                                        End If
                                                'Control Fecha desde
                                             If FECHA_DESDE <> "" Then
                                                    If Not (IsDate(FECHA_HASTA)) Then
                                                         strError = strError & "  " & " La Fecha desde No es correcta "
                                                   Else
                                                        If CDate(FECHA_HASTA) > Now And Trim(.Cells(i, C_Etiqueta) = "") Then
                                                            strError = strError & "  " & " La Fecha ES futura "
                                                        End If
                                                   End If
                                                Else
                                                    FECHA_DESDE = .Cells(i, C_Fecha_hasta)
                                                    If (Control_Documento_campos(CInt(FK_CLIENTE), .Cells(i, C_Indice), "FECHA")) <> "" Then
                                                       strError = strError & "  " & " La Fecha desde ES REQUERIDA "
                                                    End If
                                                   
                                                End If
                                                 'Control Fecha Hasta
                                                 
                                                 If Trim(.Cells(i, C_Etiqueta)) <> "" And Len(FECHA_HASTA) = 4 Then
                                                   FECHA_HASTA = "31/12/" & FECHA_HASTA
                                                End If
                    
                                                If FECHA_DESDE <> "" Then
                                                    If Not (IsDate(FECHA_HASTA)) Then
                                                         strError = strError & "  " & " La Fecha Hasta No es correcta "
                                                    Else
                                                        If Not (IsDate(FECHA_DESDE)) Then
                                                        
                                                         strError = strError & "  " & " La Fecha Hasta No es desde "
                                            
                                                        Else
                                                                If DateDiff("D", (FECHA_DESDE), (FECHA_HASTA)) < 0 Then
                                                                     strError = strError & "  " & " La Fecha Hasta es mayor que la fecha desde"
                                                                End If
                                                                 If CDate(FECHA_HASTA) > Now And Trim(.Cells(i, C_Etiqueta) = "") Then
                                                                strError = strError & "  " & " La Fecha ES futura "
                                                                End If
                                                        End If
                    
                                                    End If
                                                Else
                                                    If Trim(FECHA_DESDE) <> "" Then
                                                     
                                                        If Trim(FECHA_HASTA) = "" Then
                                                            .Cells(i, C_Fecha_hasta) = FECHA_DESDE
                                                        End If
                                                    
                                                       
                                                    End If
                                                End If
                    
                    
                    
                                        ' Control  Numero
                                                'Control Numero desde
                    
                                                If .Cells(i, C_N°_Desde) <> "" Then
                                                    If Not (IsNumeric(.Cells(i, C_N°_Desde))) Then
                                                         strError = strError & "  " & " La Numero desde No es correcta "
                                                    End If
                                                Else
                                                    .Cells(i, C_N°_Desde) = .Cells(i, C_N°_Hasta)
                                                    If (Control_Documento_campos(CInt(FK_CLIENTE), .Cells(i, C_Indice), "NUMERO")) <> "" Then
                                                       strError = strError & "  " & " La numero ES REQUERIDO "
                                                    End If
                                                End If
                    
                    
                                                 'Control Numero Hasta
                    
                                                If .Cells(i, C_N°_Hasta) <> "" Then
                                                    If Not (IsNumeric(.Cells(i, C_N°_Hasta))) Then
                                                         strError = strError & "  " & " La Numero desde No es correcta "
                                                    End If
                                                Else
                                                    If .Cells(i, C_N°_Desde) <> "" Then
                                                       strError = strError & "  " & " La Numero Hasta ES REQUERIDA "
                                                    End If
                                                End If
                    
                    
                    
                                        ' Control  Letra
                                                'Control Letra desde
                    
                                                If .Cells(i, C_Letra_Desde) = "" Then
                                                    
                                                    If (Control_Documento_campos(CInt(FK_CLIENTE), .Cells(i, C_Indice), "LETRA")) <> "" Then
                                                       strError = strError & "  " & " La letra desde ES REQUERIDO "
                                                    End If
                                                Else
                                                   .Cells(i, C_Letra_Desde) = ControlLetra(.Cells(i, C_Letra_Desde))
                                                End If
                    
                    
                                                 'Control Letra Hasta
                    
                                                If .Cells(i, C_Letra_Desde) <> "" Then
                                                   
                                                   If .Cells(i, C_Letra_Hasta) = "" Then
                                                       strError = strError & "  " & " La letra Hasta ES REQUERIDA "
                                                    Else
                                                     .Cells(i, C_Letra_Hasta) = ControlLetra(.Cells(i, C_Letra_Hasta))
                                                    End If
                                                End If
                    
                    
                                        ' Control descripcion
                    
                                            If .Cells(i, C_Descripcion) = "" Then
                                                 If (Control_Documento_campos(CInt(FK_CLIENTE), .Cells(i, C_Indice), "DESCRIPCION")) <> "" Then
                                                       strError = strError & "  " & " LA DESCRIPCION ES REQUERIDO "
                                                    End If
                                            End If
                    
                      Else
                      Debug.Print Trim(.Cells(i, C_TipoReferencia))
                      
                      End If
                      
                    
                    End If

                       
Proximo:

 If strError <> "" Then
                            .Cells(i, C_Error) = UCase(strError)
                            strError = ""
                             ErrorGeneral = True
                             
                        Else
                              If .Cells(i, C_Etiqueta) <> "" Then
                                    .Cells(i, C_Error) = "LEGA"
                               Else
                                    .Cells(i, C_Error) = "REFE"
                               End If

                        End If

        
Controlado:
                        lblCantidadRegistros.Caption = i
                        frmImportExcel.Refresh
Rem libroEx.Save
                Next

End With


           If ErrorGeneral = True Then
           Rem libroEx.SaveAs Mid(Paso, 1, Len(Paso) - 4) & " USUARIO " & MDIfrmInicio.StaInicio.Panels(2).Text & ".XLS"
           
                libroEx.Save
                
                libroEx.Close
                ApExcel.Quit
                Set hojaEx = Nothing
                Set libroEx = Nothing
                Set ApExcel = Nothing
                Rem FileCopy Paso, PasoError & NombreArchivo(Paso)
               Rem  Kill Paso
                If MsgBox("La planilla de referencia tiene errores" & vbCrLf & "   ¿Usted quiere Verificarla?   ", vbYesNo) = vbYes Then
                    Shell "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE " & Chr(34) & Paso & Chr(34), vbNormalFocus
                    Exit Sub
                Else
                Exit Sub
                
                End If
            Else
                If MsgBox("Usted Quiere continuar con la grabacion", vbInformation + vbYesNo) = vbYes Then
                With hojaEx
                    lblTitulo.Caption = "Carga: "
                     ARCHIVO_POR_LOTE = "'" & Trim(NombreArchivo(Paso)) & "'"
                     FECHA_CREACION = "'" & SysDate & "'"
                       For i = 7 To 100000
                       
                       
                       
                       
                            If .Cells(i, C_Error) = "REFE" Or .Cells(i, C_Error) = "LEGA" Then
                                FK_CLIENTE = BuscarCliente(.Cells(i, C_Caja), Cells(i, C_Cliente))
                                NRO_CAJA = .Cells(i, C_Caja)
                                If .Cells(i, C_Letra_Desde) <> "" Then
                                    LETRA_DESDE = "'" & UCase(Replace(Trim(.Cells(i, C_Letra_Desde)), "'", "´")) & "'"
                                Else
                                    LETRA_DESDE = "NULL"
                                End If
                                If .Cells(i, C_Letra_Hasta) <> "" Then
                                    LETRA_HASTA = "'" & UCase(Replace(Trim(.Cells(i, C_Letra_Hasta)), "'", "´")) & "'"
                                Else
                                    LETRA_HASTA = "NULL"
                                End If
                                
                                
                                If Not IsDate(.Cells(i, C_Fecha_desde)) Then
                                   FECHA_DESDE = "Null"
                                   Else
                                   FECHA_DESDE = FechaFormato(.Cells(i, C_Fecha_desde))
                                End If
                                
                                 
                                If Not IsDate(.Cells(i, C_Fecha_hasta)) Then
                                   FECHA_HASTA = "Null"
                                   Else
                                   FECHA_HASTA = FechaFormato(.Cells(i, C_Fecha_hasta))
                                End If
                                ' DECRIPCION
                                If Trim(.Cells(i, C_Descripcion)) <> "" Then
                                    Descripcion = "'" & UCase(Trim(.Cells(i, C_Descripcion))) & "'"
                                Else
                                    Descripcion = "NULL"
                                End If
                                Rem Documento
                                If (.Cells(i, C_N°_Desde)) <> "" Then
                                    NRO_DESDE = .Cells(i, C_N°_Desde)
                                Else
                                    NRO_DESDE = "NULL"
                                End If
                                If (.Cells(i, C_N°_Hasta)) <> "" Then
                                    NRO_HASTA = .Cells(i, C_N°_Hasta)
                                Else
                                    NRO_HASTA = "NULL"
                                End If
                                
                               
                                FK_INDICES = Buscar_ID_Indice(.Cells(i, C_Indice), CInt(FK_CLIENTE))
                                Cod_Indice = "'" & BuscarIndice(CInt(FK_CLIENTE), .Cells(i, C_Indice)) & "'"
                                
                                If .Cells(i, C_Error) = "REFE" Then
                                
                                If Trim(Cod_Indice) = "'0'" Then
                                    MsgBox "Error en el indice "
                                    Exit Sub
                                End If
                                
                                    ID_referencia = InsertarReferencias(CInt(FK_CLIENTE), CLng(NRO_CAJA), CStr(Cod_Indice), CStr(Descripcion) _
                                    , CStr(FECHA_DESDE), CStr(FECHA_HASTA), CStr(NRO_DESDE), CStr(NRO_HASTA), CStr(LETRA_DESDE), CStr(LETRA_HASTA), "Null" _
                                    , CStr(ARCHIVO_POR_LOTE), 0, ARCHIVO_POR_LOTE, MDIfrmInicio.StaInicio.Panels(2).Text)
                                    .Cells(i, C_Error) = "SE REGISTRO LA REFENCIA CON EL ID  " & ID_referencia
                                
                                    .Cells(i, C_Error) = "SE REGISTRO " & ID_referencia
                                End If
                                    If .Cells(i, C_Error) = "LEGA" Then
                                        Sql = " Update LEGAJOS"
                                        Sql = Sql & " SET FK_INDICES =" & FK_INDICES
                                        Sql = Sql & " , COD_INDICE =" & Cod_Indice
                                        Sql = Sql & " , NRO_DESDE =" & NRO_DESDE
                                        Sql = Sql & " , NRO_HASTA =" & NRO_HASTA
                                        Sql = Sql & " , LETRA_DESDE =" & LETRA_DESDE
                                        Sql = Sql & " , LETRA_HASTA =" & LETRA_HASTA
                                        Sql = Sql & " , FECHA_DESDE =" & FECHA_DESDE
                                        Sql = Sql & " , FECHA_HASTA =" & FECHA_HASTA
                                        Sql = Sql & " , FECHA_CREACION  =" & SysDate
                                        Sql = Sql & " , FK_PERSONAL_CREACION = " & MDIfrmInicio.StaInicio.Panels(2).Text
                                        Sql = Sql & " , NRO_CAJA =" & NRO_CAJA
                                        Sql = Sql & " , COD_CLIENTE =" & FK_CLIENTE
                                        Sql = Sql & " , DESCRIPCION =" & Descripcion
                                        Sql = Sql & " , COD_ESTADO =2"
                                        Sql = Sql & " , PASOARCHIVO= " & ARCHIVO_POR_LOTE
                                        If Trim(.Cells(i, C_Lote)) <> "" Then
                                         If IsNumeric(.Cells(i, C_Lote)) Then
                                            Sql = Sql & " , LOTE= " & .Cells(i, C_Lote)
                                         End If
                                        
                                        End If
                                        Sql = Sql & " Where ID_LEGAJO =" & .Cells(i, C_Etiqueta)
                                        Sql = Sql & " AND NRO_CAJA IS NULL"
                                         R = ExecutarSql(Sql)
                                        If R = 0 Then
                                            .Cells(i, C_Error) = "NO SE REGISTRO"
                                            MsgBox "NO SE REGISTRO EL LEGAJO " & .Cells(i, C_Etiqueta)
                                        Else
                                            .Cells(i, C_Error) = "SE " & R & " LEGAJOS "
                                        End If
                                        
                                    End If
                                    lblCantidadRegistros.Caption = i
                                Else
                                Exit For
                                End If
                            Next
                    End With
         MsgBox " Se actualizaron " & i
         End If
  
 End If
libroEx.Save
libroEx.Close
ApExcel.Quit
Set hojaEx = Nothing
Set libroEx = Nothing
Set ApExcel = Nothing

 FileCopy Paso, Mid(Paso, 1, Len(Paso) - 4) & " procesado por " & MDIfrmInicio.StaInicio.Panels(2) & " .xls"
 Kill Paso

clienteCajasChicas = 0
MsgBox "Terminado", vbInformation
 Exit Sub
salir:
MsgBox " Error " & Err.Description
clienteCajasChicas = 0
End Sub

Public Function ControlLetra(lETRA As String) As String
    Dim Letracontrol As String
    

Letracontrol = Trim(UCase(lETRA))

Letracontrol = Replace(Letracontrol, "*", " ")
Letracontrol = Replace(Letracontrol, ".", "")
Letracontrol = Replace(Letracontrol, "(", "")
Letracontrol = Replace(Letracontrol, ",", " ")
Letracontrol = Replace(Letracontrol, "/", " ")

Letracontrol = Replace(Letracontrol, "  ", " ")
Letracontrol = Replace(Letracontrol, "   ", " ")
Letracontrol = Replace(Letracontrol, "    ", " ")
Letracontrol = Replace(Letracontrol, "     ", " ")
Letracontrol = Replace(Letracontrol, "Á", "A")
Letracontrol = Replace(Letracontrol, "É", "E")
Letracontrol = Replace(Letracontrol, "Í", "I")
Letracontrol = Replace(Letracontrol, "Ó", "O")
Letracontrol = Replace(Letracontrol, "Ú", "U")

Letracontrol = Trim(Letracontrol)

ControlLetra = Letracontrol

End Function

Public Function BuscarCliente(CAJAS As Long, Cliente_Excel As String) As Long

Dim Sql As String

If CAJAS < 100000 Then
    
    If IsNumeric(Cliente_Excel) Then
        BuscarCliente = Cliente_Excel
    Else
        BuscarCliente = InputBox("Ingrese el cliente para la caja " & CAJAS, "Cajas sin asignar", clienteCajasChicas)
        clienteCajasChicas = BuscarCliente
    End If
    
    
    
    Sql = " SELECT     ID_CAJA, FK_CLIENTE"
    Sql = Sql & " From cajas"
    Sql = Sql & "  Where NRO_CAJA = " & CAJAS
    Sql = Sql & " AND FK_CLIENTE =   " & BuscarCliente
      
Else
    Sql = " SELECT     ID_CAJA, FK_CLIENTE"
    Sql = Sql & " From cajas"
    Sql = Sql & "  Where ID_CAJA = " & CAJAS
End If





Dim rs As New ADODB.Recordset

    rs.Open Sql, ConActiva, 0, 1
    
    If rs.EOF Then
    BuscarCliente = 0
    Else
    If IsNull(rs!FK_CLIENTE) Then
        BuscarCliente = 0
    Else
        BuscarCliente = rs!FK_CLIENTE
    End If
    
    End If
    



End Function

Public Function ControlCorreo(FK_CLIENTE As Integer, Caja As Long) As String

Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = " SELECT     CLIENTEUSUARIO.CORREO"
Sql = Sql & " FROM         REMITOS_CUERPO INNER JOIN "
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
Sql = Sql & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
Sql = Sql & " WHERE     (REMITOS_CUERPO.TIPO = 0)"
Sql = Sql & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
Sql = Sql & " AND  REMITOS_DETALLE.DESDE =  " & Caja
Sql = Sql & " AND REMITOS_CUERPO.ID_CLIENTE = " & FK_CLIENTE

rs.Open Sql, ConActiva


If rs.EOF Then
    ControlCorreo = ""
Else
If IsNull(rs!correo) Then
 ControlCorreo = ""
 Else
 ControlCorreo = Trim(rs!correo)
 End If
 
End If

End Function

Public Function Doc_Usuario(FK_CLIENTE As Integer, Caja As Long) As Long

Dim Sql As String
Dim rs As New ADODB.Recordset
    Doc_Usuario = 0

Sql = " SELECT      INDICES.ID_CODIGO_DOCUMENTO"
Sql = Sql & vbCrLf & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & vbCrLf & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
Sql = Sql & vbCrLf & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO INNER JOIN"
Sql = Sql & vbCrLf & " INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE"
Sql = Sql & vbCrLf & " WHERE     (REMITOS_CUERPO.TIPO = 0) "
Sql = Sql & vbCrLf & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) AND "
Sql = Sql & vbCrLf & " REMITOS_DETALLE.DESDE = " & Caja
Sql = Sql & vbCrLf & " AND REMITOS_CUERPO.ID_CLIENTE = " & FK_CLIENTE

rs.Open Sql, ConActiva


If rs.EOF Then
    Doc_Usuario = 0
Else
If IsNull(rs!ID_CODIGO_DOCUMENTO) Then
 Doc_Usuario = 0
 Else
 Doc_Usuario = Trim(rs!ID_CODIGO_DOCUMENTO)
 End If
 
End If




End Function

Public Sub BorrarReferenciaAdministradaPorCLiente(Cliente As Integer, Caja As Long)

Dim Sql As String
If chkBorrarReferencia.value = 1 Then


    Sql = " DELETE FROM REFERENCIAS"
    Sql = Sql & " WHERE     (DESCRIPCION LIKE '%Referencia administrada por el cliente%')"
    Sql = Sql & " AND NRO_CAJA = " & Caja
    Sql = Sql & " AND COD_CLIENTE = " & Cliente
    ExecutarSql Sql
End If
End Sub

Public Sub BorrarReferencias(Cliente As Integer, Caja As Long)
Dim Sql As String
    Sql = " DELETE FROM REFERENCIAS"
    Sql = Sql & " WHERE    "
    Sql = Sql & " NRO_CAJA = " & Caja
    Sql = Sql & " AND COD_CLIENTE = " & Cliente
    ExecutarSql Sql
End Sub

Public Function Indice_Secundario_Disco(Indice As Integer) As String
    Dim DATO As String
        Select Case Indice
        Case 1
            DATO = "Actas de seguridad - libros de actas"
        Case 2
            DATO = "Partes de entrada de mercaderias - Devoluciones al cdc - Decomisos - Hojas de Ruta - Devoluciones al proveedor"
        Case 3
            DATO = "Libros IVA compras - ventas "
        Case 4
            DATO = "Notas de Creditos - Notas de Debitos - Arqueos de Caja - Arqueos de tesoreria"
        Case 5
            DATO = "Planillas de inventario - Ajustes por inventario - Planillas de inventario primer conteo "
        Case 6
            DATO = "Remitos manuales de decomisos - devoluciones - transferencias"
        Case 7
            DATO = "Transferencias del cdc al local - Transferencias de mercaderias"
        Case 8
            DATO = "Planillas de ventas"
        End Select
        Indice_Secundario_Disco = DATO
End Function

Public Function AcutalizarImagenRemito(ID_imagen As Long, Filtro As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim Sql     As String
        Sql = " SELECT NRO_REMITO, NRO_REM_PROV "
        Sql = Sql & " From REMITOS_CUERPO "
        Sql = Sql & " Where " & Filtro
        rs.Open Sql, strConBasa
        If Not rs.EOF Then
            Sql = "  Update DOCUMENTOS_DIGITALES"
            Sql = Sql & " SET NRO_DESDE =" & rs!NRO_REMITO
            Sql = Sql & " , NRO_HASTA =" & rs!NRO_REMITO
            Sql = Sql & ", LETRA_DESDE ='" & Trim(rs!NRO_REM_PROV) & "'"
            Sql = Sql & ", LETRA_HASTA ='" & Trim(rs!NRO_REM_PROV) & "'"
            Sql = Sql & " Where ID = " & ID_imagen
        End If
        ExecutarSql Sql
End Function

Public Function FECHA_DESDE_SOLUCION(Dia, Mes, Año) As String

        FECHA_DESDE_SOLUCION = ""
        If IsNull(Dia) Or Not IsNumeric(Dia) Then
            Dia = ""
        Else
            If Dia > 31 Then
                Dia = ""
                Mes = ""
                Año = ""
            End If
        End If
        
        If IsNull(Mes) Or Not IsNumeric(Mes) Then
            Mes = ""
        Else
            If Mes = 20 Then
               Dia = ""
               Mes = ""
            End If
            
            
            If Mes > 12 And Mes <> "" Then
               Dia = ""
               Mes = ""
               Año = ""
            End If
        
        End If
        
        If IsNull(Año) Or Not IsNumeric(Año) Then
           Año = ""
        End If
        


            If Dia = "" And Mes = "" And Año <> "" Then
                   FECHA_DESDE_SOLUCION = Format("01/01/" & Año, "DD/MM/YYYY")
                   Exit Function
            End If
                            
            If Dia = "" And Mes <> "" And Año <> "" Then
                   FECHA_DESDE_SOLUCION = Format("01/" & Format(Mes, "00") & "/" & Año, "DD/MM/YYYY")
                   Exit Function
            End If
               
            If Dia <> "" And Mes <> "" And Año <> "" Then
                FECHA_DESDE_SOLUCION = Format(Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Año, "00"), "DD/MM/YYYY")
            Else
                FECHA_DESDE_SOLUCION = ""
            End If



End Function

Public Function FECHA_HASTA_SOLUCION(Dia, Mes, Año) As String

        FECHA_HASTA_SOLUCION = ""
        If IsNull(Dia) Or Not IsNumeric(Dia) Then
            Dia = ""
        Else
         If Dia > 31 Then
            Dia = ""
            Mes = ""
            Año = ""
         End If
         
        End If
        
        If IsNull(Mes) Or Not IsNumeric(Mes) Then
            Mes = ""
        Else
            If Mes = 20 Then
               Dia = ""
               Mes = ""
            End If
            If Mes > 12 And Mes <> "" Then
               Dia = ""
               Mes = ""
               Año = ""
            End If
        
        End If
        
        If IsNull(Año) Or Not IsNumeric(Año) Then
           Año = ""
        End If
        


            If Dia = "" And Mes = "" And Año <> "" Then
                   FECHA_HASTA_SOLUCION = Format("31/12/" & Año, "DD/MM/YYYY")
                   Exit Function
            End If
                            
            If Dia = "" And Mes <> "" And Año <> "" Then
                   Select Case Mes
                   Case 1, 3, 5, 7, 8, 10, 12
                        Dia = "31"
                   Case 4, 6, 9, 11
                        Dia = "30"
                   Case 2
                        Dia = "28"
                   End Select
                   
                   FECHA_HASTA_SOLUCION = Format(Dia & "/" & Format(Mes, "00") & "/" & Año, "DD/MM/YYYY")
                   Exit Function
            End If
               
            If Dia <> "" And Mes <> "" And Año <> "" Then
                FECHA_HASTA_SOLUCION = Format(Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Año, "00"), "DD/MM/YYYY")
            Else
                FECHA_HASTA_SOLUCION = ""
            End If



End Function

Public Sub LeerExcelRecetas(Paso As String, MesAño As String, Quincena As String, Indice_Doc As Integer, NOMBRE_PLANILLA As String)
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet

    Dim i As Integer
    Dim C_Caja As Integer
    Dim Caja_anterior As Long
    Dim C_Nombre As Integer
    Dim C_Cod_Pami As Integer
    Dim C_Cod_Flk As Integer
    Dim C_Lotes As Integer
    Dim Sql As String
    
    Dim Caja  As String
    Dim Indice As String
    Dim Descripcion As String
    Dim NRO_DESDE As String
    Dim NRO_HASTA As String
    Dim FECHA_HASTA As String
    Dim FECHA_DESDE As String
    Dim LETRA_DESDE As String
    Dim LETRA_HASTA As String
   
    C_Caja = 1
    C_Cod_Flk = 2
    C_Nombre = 3
    C_Cod_Pami = 4
    C_Lotes = 5
    

    
    
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Rem Set libroEx = Excel.Workbooks.Open("\\222.15.19.251\basa\Administracion\Referencias\" & "Planilla Modelo.xls", , True)
      Set libroEx = Excel.Workbooks.Open(Paso, True)
    
    Set hojaEx = libroEx.Worksheets.Item(1)
            
 For i = 1 To 9000
'    MsgBox hojaEx.Cells(I, C_Caja)
'    MsgBox hojaEx.Cells(I, C_Nombre)
'    MsgBox hojaEx.Cells(I, C_Cod_Pami)
    
    If IsNumeric(hojaEx.Cells(i, C_Cod_Flk)) And Trim(hojaEx.Cells(i, C_Cod_Flk)) <> "" Then
    
    If Trim(hojaEx.Cells(i, C_Caja)) <> "" Then
    
    Caja = hojaEx.Cells(i, C_Caja)
    
    If Not IsNumeric(Caja) Then
     MsgBox "error en planilla " & Paso
    End If
    
    
    If Caja_anterior <> Caja Then
    Caja_anterior = Caja
    End If
    End If
    
    
    Indice = Indice_Doc
 
    
    Descripcion = CStr(hojaEx.Cells(i, C_Nombre))
    
    If Trim(CStr(hojaEx.Cells(i, C_Cod_Pami))) <> "" Then
         
            If Not IsNumeric(Trim(CStr(hojaEx.Cells(i, C_Cod_Pami)))) Then
                 NRO_DESDE = 0
                 NRO_HASTA = 0
            
            
            Else
                NRO_DESDE = hojaEx.Cells(i, C_Cod_Pami)
                NRO_HASTA = hojaEx.Cells(i, C_Cod_Pami)
             End If
    Else
     NRO_DESDE = 0
                 NRO_HASTA = 0
    
    End If
    
    If Quincena = 1 Then
        FECHA_HASTA = "01/" & Mid(MesAño, 1, 2) & "/" & Mid(MesAño, 3)
        FECHA_DESDE = "15/" & Mid(MesAño, 1, 2) & "/" & Mid(MesAño, 3)
    End If
    If Quincena = 2 Then
        If Mid(MesAño, 1, 2) = "02" Then
           FECHA_HASTA = "16/" & Mid(MesAño, 1, 2) & "/" & Mid(MesAño, 3)
           FECHA_DESDE = "28/" & Mid(MesAño, 1, 2) & "/" & Mid(MesAño, 3)
        Else
           FECHA_HASTA = "16/" & Mid(MesAño, 1, 2) & "/" & Mid(MesAño, 3)
           FECHA_DESDE = "30/" & Mid(MesAño, 1, 2) & "/" & Mid(MesAño, 3)
        End If
    End If
    
    
    If Trim(hojaEx.Cells(i, C_Lotes)) = "" Then
        LETRA_DESDE = "SIN LOTE"
        LETRA_HASTA = "SIN LOTE"
    Else
        LETRA_DESDE = Trim(hojaEx.Cells(i, C_Lotes))
        LETRA_HASTA = Trim(hojaEx.Cells(i, C_Lotes))
    End If
    
    
    
        Sql = "Insert Into TEM_OSEP_RECETAS("
        Sql = Sql & vbCrLf & " Caja "
        Sql = Sql & vbCrLf & ", Indice"
        Sql = Sql & vbCrLf & ", Descripcion"
        Sql = Sql & vbCrLf & ", NRO_DESDE"
        Sql = Sql & vbCrLf & ", NRO_HASTA"
        Sql = Sql & vbCrLf & ", FECHA_HASTA"
        Sql = Sql & vbCrLf & ", FECHA_DESDE"
        Sql = Sql & vbCrLf & ", NOMBRE_PLANILLA"
        Sql = Sql & vbCrLf & ",LETRA_DESDE"
        Sql = Sql & vbCrLf & ",LETRA_HASTA"
        Sql = Sql & vbCrLf & ",ARCHIVO"
        Sql = Sql & vbCrLf & ")"
        Sql = Sql & vbCrLf & " VALUES("
        Sql = Sql & vbCrLf & Caja_anterior
        Sql = Sql & vbCrLf & "," & Indice
        Sql = Sql & vbCrLf & ",'" & Trim(Descripcion) & "'"
        Sql = Sql & vbCrLf & "," & NRO_DESDE
        Sql = Sql & vbCrLf & "," & NRO_HASTA
        Sql = Sql & vbCrLf & ",'" & FECHA_HASTA & "'"
        Sql = Sql & vbCrLf & ",'" & FECHA_DESDE & "'"
        Sql = Sql & vbCrLf & ",'" & NOMBRE_PLANILLA & "'"
        Sql = Sql & vbCrLf & ",'" & LETRA_DESDE & "'"
        Sql = Sql & vbCrLf & ",'" & LETRA_HASTA & "'"
         Sql = Sql & vbCrLf & ",'" & txtQuincenaMesAño.Text & "'"
        Sql = Sql & vbCrLf & ")"
        ExecutarSql Sql
    End If
 
 Next


Set hojaEx = Nothing
Set libroEx = Nothing
Set ApExcel = Nothing
Exit Sub

End Sub

Public Sub CrearPlanilla()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim DATO As String
    Sql = " SELECT     ID, CAJA, INDICE, DESCRIPCION, NRO_DESDE, NRO_HASTA, convert(char, FECHA_HASTA,103) as FECHA_HASTA ,convert(char, FECHA_DESDE,103) as FECHA_DESDE , NOMBRE_PLANILLA, LETRA_DESDE, LETRA_HASTA, ARCHIVO"
    Sql = Sql & " From basasql.dbo.TEM_OSEP_RECETAS"
    Sql = Sql & " WHERE     (ARCHIVO = '" & txtQuincenaMesAño.Text & "')"
    Sql = Sql & " ORDER BY ID "

    rs.Open Sql, strConBasa
    Do While Not rs.EOF
        DATO = DATO & vbCrLf & Trim(rs!Caja) & vbTab & Trim(rs!Indice) & vbTab & Trim(rs!Descripcion) & vbTab & Trim(rs!NRO_DESDE) & vbTab & Trim(rs!NRO_HASTA) & vbTab & Trim(rs!FECHA_HASTA) & vbTab & Trim(rs!FECHA_DESDE) & vbTab & Trim(rs!LETRA_DESDE) & vbTab & Trim(rs!LETRA_HASTA) & vbTab & Trim(rs!Archivo) & vbTab & Trim(rs!NOMBRE_PLANILLA)
        rs.MoveNext
    Loop
    
    
    
Clipboard.GetData
Clipboard.Clear
Clipboard.SetText DATO
MsgBox ("Datos copiados")
    
End Sub
