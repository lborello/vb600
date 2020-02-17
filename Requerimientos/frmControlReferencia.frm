VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControlReferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Referencia"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscarRemito 
      Caption         =   "Buscar Remito"
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
      Left            =   3300
      TabIndex        =   5
      Top             =   6240
      Width           =   1515
   End
   Begin VB.TextBox txtRemito 
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
      Left            =   1500
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
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
      Left            =   9960
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid grdControlReferencia 
      Height          =   5655
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar Remito :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6300
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Control de referencia por remito"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   60
      Width           =   11115
   End
End
Attribute VB_Name = "frmControlReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub cmdActualizar_Click()
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Dim sql As String
        Dim Clave As String
       
            sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, ID_CLIENTE,"
            sql = sql & vbCrLf & " CANTIDAD, DESC_CONTROL_REF,"
            sql = sql & vbCrLf & "  CONTROL_REFERENCIA"
            sql = sql & vbCrLf & "  From REMITOS_CUERPO"
            sql = sql & vbCrLf & "  WHERE (CONTROL_REFERENCIA IS NULL) AND (TIPO = 0) AND"
            sql = sql & vbCrLf & "    (NRO_REMITO > 56000) ORDER BY NRO_REMITO"
            rs.Open sql, ConActiva, adOpenDynamic, adLockOptimistic
            DATOSGRILLA grdControlReferencia, rs
            
            grdControlReferencia.Columns.Item(0).Alignment = dbgCenter
            grdControlReferencia.Columns.Item(0).Locked = True
            
            grdControlReferencia.Columns.Item(1).Locked = True
            grdControlReferencia.Columns.Item(1).Alignment = dbgCenter
            
            grdControlReferencia.Columns.Item(2).Locked = True
            grdControlReferencia.Columns.Item(2).Alignment = dbgCenter
            
            grdControlReferencia.Columns.Item(3).Locked = True
            grdControlReferencia.Columns.Item(3).Alignment = dbgCenter
            
            grdControlReferencia.Columns.Item(4).Locked = True
            grdControlReferencia.Columns.Item(4).Alignment = dbgCenter
            
            grdControlReferencia.Refresh
    
End Sub

Public Sub DATOSGRILLA(Grilla As DataGrid, rs As ADODB.Recordset)
Grilla.ClearFields
Grilla.ClearSelCols
Grilla.ScrollBars = dbgAutomatic
Dim i As Integer
For i = 0 To rs.Fields.Count - 1
    
    Debug.Print rs.Fields.Item(i).Name & "  " & rs.Fields.Item(i).Type
    
    Grilla.Columns.Add i
    Grilla.Columns.Item(i).DataField = rs.Fields(i).Name
    Grilla.Columns.Item(i).Caption = rs.Fields(i).Name
    Select Case rs.Fields.Item(i).Type
    Case "131" ' NUMERO
        Grilla.Columns.Item(i).Width = 500
    Case "200" 'TEXT
        Grilla.Columns.Item(i).Width = 1500
    Case "135" 'FECHA
        Grilla.Columns.Item(i).Width = 700
    End Select
    
Next

Set Grilla.DataSource = rs.DataSource
Grilla.Refresh


End Sub

Private Sub cmdBuscarRemito_Click()
 Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Dim sql As String
        Dim Clave As String
       
            sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, ID_CLIENTE,"
            sql = sql & vbCrLf & " CANTIDAD, DESC_CONTROL_REF,"
            sql = sql & vbCrLf & " CONTROL_REFERENCIA"
            sql = sql & vbCrLf & " From REMITOS_CUERPO"
            sql = sql & vbCrLf & " WHERE (TIPO = 0) AND"
            sql = sql & vbCrLf & " NRO_REMITO = " & txtRemito.Text
            sql = sql & vbCrLf & " ORDER BY NRO_REMITO"
            rs.Open sql, ConActiva, adOpenDynamic, adLockOptimistic
            DATOSGRILLA grdControlReferencia, rs
            
            grdControlReferencia.Columns.Item(0).Alignment = dbgCenter
            grdControlReferencia.Columns.Item(0).Locked = True
            
            grdControlReferencia.Columns.Item(1).Locked = True
            grdControlReferencia.Columns.Item(1).Alignment = dbgCenter
            
            grdControlReferencia.Columns.Item(2).Locked = True
            grdControlReferencia.Columns.Item(2).Alignment = dbgCenter
            
            grdControlReferencia.Columns.Item(3).Locked = True
            grdControlReferencia.Columns.Item(3).Alignment = dbgCenter
            
            grdControlReferencia.Columns.Item(4).Locked = True
            grdControlReferencia.Columns.Item(4).Alignment = dbgCenter
            
            
            grdControlReferencia.Columns.Item(5).Locked = True
            grdControlReferencia.Columns.Item(5).Alignment = dbgCenter
            
            
            grdControlReferencia.Columns.Item(6).Locked = True
            grdControlReferencia.Columns.Item(6).Alignment = dbgCenter
            
           
            grdControlReferencia.Refresh
End Sub

