VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmHojaRuta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoja de Ruta"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12765
   Begin TabDlg.SSTab SSTab1 
      Height          =   8835
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   15584
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmHojaRuta.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtFiltro"
      Tab(0).Control(1)=   "cboHuta"
      Tab(0).Control(2)=   "cmdActualizar"
      Tab(0).Control(3)=   "grdRutaCuerpo"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Hoja de ruta"
      TabPicture(1)   =   "frmHojaRuta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdHojarutadetalle"
      Tab(1).Control(1)=   "cmdAceptar"
      Tab(1).Control(2)=   "txtRequerimiento"
      Tab(1).Control(3)=   "cmdInsertarRequerimiento"
      Tab(1).Control(4)=   "txtID_Hoja"
      Tab(1).Control(5)=   "txtFecha"
      Tab(1).Control(6)=   "cmdImpirmir"
      Tab(1).Control(7)=   "cmdRefrescar"
      Tab(1).Control(8)=   "ctlPersonal"
      Tab(1).Control(9)=   "grdDetalleHojaRuta"
      Tab(1).Control(10)=   "Label1(0)"
      Tab(1).Control(11)=   "Label1(1)"
      Tab(1).Control(12)=   "Label1(2)"
      Tab(1).Control(13)=   "Label1(3)"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmHojaRuta.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdHojarutadetalle 
         Caption         =   "Detalle de hoja de ruta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   18
         Top             =   8280
         Width           =   2055
      End
      Begin VB.TextBox txtFiltro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70560
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboHuta 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmHojaRuta.frx":0054
         Left            =   -74760
         List            =   "frmHojaRuta.frx":005E
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   960
         Width           =   3975
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "actualizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67680
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid grdRutaCuerpo 
         Height          =   5595
         Left            =   -74760
         TabIndex        =   14
         Top             =   1500
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9869
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Height          =   330
         Left            =   -63840
         TabIndex        =   9
         Top             =   8280
         Width           =   1200
      End
      Begin VB.TextBox txtRequerimiento 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73020
         TabIndex        =   6
         Top             =   1800
         Width           =   4035
      End
      Begin VB.CommandButton cmdInsertarRequerimiento 
         Caption         =   "Insertar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68820
         TabIndex        =   5
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox txtID_Hoja 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73020
         TabIndex        =   4
         Top             =   540
         Width           =   4035
      End
      Begin VB.TextBox txtFecha 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73020
         TabIndex        =   3
         Top             =   960
         Width           =   4035
      End
      Begin VB.CommandButton cmdImpirmir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -65160
         TabIndex        =   2
         Top             =   8280
         Width           =   1200
      End
      Begin VB.CommandButton cmdRefrescar 
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
         Height          =   330
         Left            =   -66480
         TabIndex        =   1
         Top             =   8280
         Width           =   1200
      End
      Begin Controles.cltGenerico ctlPersonal 
         Height          =   315
         Left            =   -73020
         TabIndex        =   7
         Top             =   1380
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
      End
      Begin MSDataGridLib.DataGrid grdDetalleHojaRuta 
         Height          =   5835
         Left            =   -74880
         TabIndex        =   8
         Top             =   2280
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   10292
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Caption         =   "Detalle de hoja de ruta"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "COD_CLIENTE"
            Caption         =   "Cliente"
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
            DataField       =   "COD_REQUERIMIENTO"
            Caption         =   "Requer"
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
         BeginProperty Column02 
            DataField       =   "DETALLE_TIPO"
            Caption         =   "Tipo"
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
         BeginProperty Column03 
            DataField       =   "ORDEN"
            Caption         =   "Orden"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "KILOMETROS"
            Caption         =   "KM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "DETALLE"
            Caption         =   "Detalle"
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
               Locked          =   -1  'True
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   3390,236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3960
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Hoja de ruta:"
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
         Index           =   0
         Left            =   -74820
         TabIndex        =   13
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
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
         Index           =   1
         Left            =   -74820
         TabIndex        =   12
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Personal:"
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
         Index           =   2
         Left            =   -74820
         TabIndex        =   11
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Insertar Requerimiento:"
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
         Index           =   3
         Left            =   -74820
         TabIndex        =   10
         Top             =   1860
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmHojaRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    UpdateHojaRuta
End Sub

Private Sub cmdActualizar_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sql = " SELECT     TOP 100 ID_HOJA_RUTA, FECHA, CANTIDADKILOMETROS, COD_PERSONAL, ESTADO"
    sql = sql & " From HOJA_RUTA_CUERPO"
    sql = sql & " ORDER BY ID_HOJA_RUTA DESC"
    rs.Open sql, ConActiva, adOpenDynami, adLockReadOnly
    Set grdRutaCuerpo.DataSource = rs.DataSource
    grdRutaCuerpo.Refresh
End Sub

Private Sub cmdHojarutadetalle_Click()
        Dim sql As String
        sql = " SELECT * "
        sql = sql & " FROM  V_HOJA_RUTA_DETALLE "
        sql = sql & " WHERE  ID_HOJA_RUTA=" & txtID_Hoja.Text
        sql = sql & "  ORDER BY V_HOJA_RUTA_DETALLE.COD_CLIENTE , IDREQUERIMIENTO  "
        frmReportes.ImprimirReporte PasoReportes & "RPT_Hoja_Ruta_Detalle.rpt", sql, True
End Sub

Private Sub cmdImpirmir_Click()
    Dim sql As String
        sql = " SELECT * "
        sql = sql & " FROM  V_HOJARUTA"
        sql = sql & " WHERE  ID_HOJA_RUTA=" & txtID_Hoja.Text
        sql = sql & "  ORDER BY V_HOJARUTA.COD_CLIENTE"
        Rem ojo esta ordenada la vista
        Rem  sql = sql & " ORDER BY ORDEN"
        frmReportes.ImprimirReporte PasoReportes & "rptHojaRuta2.rpt", sql, True

End Sub

Private Sub cmdInsertarRequerimiento_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = " SELECT REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.DESCRIPCION,"
    sql = sql & " REQUERIMIENTO.ID_CLIENTE,TIPOREQUERIMIENTO.DESCRIPCION AS DETALLE_TIPO_REQUERIMIENTO"
    sql = sql & " From Requerimiento, Tiporequerimiento"
    sql = sql & " Where Requerimiento.IDTIPOREQUERIMIENTO = Tiporequerimiento.IDTIPOREQUERIMIENTO"
    sql = sql & " AND IDREQUERIMIENTO IN (" & txtRequerimiento & " )"
    sql = sql & " ORDER BY ID_CLIENTE "
    rs.Open sql, ConActiva, 0, 1
    Do While Not rs.EOF
        If Not IsNull(rs!DESCRIPCION) Then
            DETALLE = "'" & Trim(rs!DESCRIPCION) & "'"
        Else
            DETALLE = "Null"
        End If
            sql = " INSERT INTO HOJA_RUTA_DETALLE"
            sql = sql & " (COD_HOJA_RUTA, COD_REQUERIMIENTO, "
            sql = sql & " DETALLE,COD_CLIENTE,DETALLE_TIPO_REQUERIMIENTO)"
            sql = sql & " VALUES "
            sql = sql & "(" & txtID_Hoja.Text & "," & rs!IDREQUERIMIENTO & ","
            sql = sql & DETALLE & "," & rs!ID_CLIENTE & ",'" & Trim(rs!DETALLE_TIPO_REQUERIMIENTO) & "')"
            ExecutarSql sql
            rs.MoveNext
    Loop
    ActualizarGrilla txtID_Hoja.Text
End Sub

Private Sub Command1_Click()
ActualizarGrilla 1



End Sub

Public Sub AgregarHojaRuta(FiltroRequerimiento As String)
    Dim sql As String
    Dim ID_Hoja As Long
    Dim DETALLE As String
    ID_Hoja = MaxRuta
    On Error GoTo salir
    
    sql = " INSERT INTO HOJA_RUTA_CUERPO "
    sql = sql & "( FECHA, CANTIDADKILOMETROS,"
    sql = sql & " COD_PERSONAL)"
    sql = sql & " VALUES "
    sql = sql & " (" & SysDate & ",0,"
    sql = sql & 99 & ")"
    ExecutarSql sql
    txtID_Hoja.Text = ID_Hoja
    txtFecha.Text = SysDate_DD_MM_YYYY
    
    
    Dim rs As New ADODB.Recordset
    
    sql = " SELECT REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.DESCRIPCION,"
    sql = sql & "  REQUERIMIENTO.ID_CLIENTE,TIPOREQUERIMIENTO.DESCRIPCION AS DETALLE_TIPO_REQUERIMIENTO"
    sql = sql & " From Requerimiento, Tiporequerimiento"
    sql = sql & "  Where Requerimiento.IDTIPOREQUERIMIENTO = Tiporequerimiento.IDTIPOREQUERIMIENTO"
    sql = sql & "   AND IDREQUERIMIENTO IN (" & FiltroRequerimiento & " )"
   
    rs.Open sql, strConBasa, 0, 1
    Do While Not rs.EOF
    If Not IsNull(rs!DESCRIPCION) Then
        DETALLE = "'" & Trim(rs!DESCRIPCION) & "'"
    Else
        DETALLE = "Null"
    End If
        sql = " INSERT INTO HOJA_RUTA_DETALLE"
        sql = sql & " (COD_HOJA_RUTA, COD_REQUERIMIENTO, "
        sql = sql & " DETALLE,COD_CLIENTE,DETALLE_TIPO_REQUERIMIENTO)"
        sql = sql & " VALUES "
        sql = sql & "(" & ID_Hoja & "," & rs!IDREQUERIMIENTO & ","
        sql = sql & DETALLE & "," & rs!ID_CLIENTE & ",'" & Trim(rs!DETALLE_TIPO_REQUERIMIENTO) & "')"
        ExecutarSql sql
        rs.MoveNext
    Loop
    ActualizarGrilla ID_Hoja
    Exit Sub
salir:
    MsgBox Err.Description
    

End Sub

Public Function MaxRuta() As Long
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = " SELECT MAX(ID_HOJA_RUTA) AS MaxRuta From HOJA_RUTA_CUERPO"
    rs.Open sql, ConActiva, 0, 1
    MaxRuta = rs!MaxRuta + 1
End Function

Public Sub ActualizarGrilla(ID_RUTA As Long)
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Dim sql As String
        sql = " SELECT COD_CLIENTE , COD_REQUERIMIENTO ,"
        sql = sql & vbCrLf & " DETALLE_TIPO_REQUERIMIENTO AS DETALLE_TIPO , ORDEN , KILOMETROS , DETALLE,ID"
        sql = sql & vbCrLf & " From HOJA_RUTA_DETALLE"
        sql = sql & vbCrLf & " Where COD_HOJA_RUTA = " & ID_RUTA
        sql = sql & vbCrLf & " ORDER BY ORDEN, COD_CLIENTE"
        rs.Open sql, ConActiva, adOpenDynamic, adLockOptimistic
        Set grdDetalleHojaRuta.DataSource = rs.DataSource
        grdDetalleHojaRuta.Refresh

End Sub


Private Sub cmdRefrescar_Click()
    ActualizarGrilla txtID_Hoja.Text
     AcualizarRequerimiento txtID_Hoja.Text
End Sub

Private Sub Form_Load()
    ctlPersonal.TipoControl = PERSONAL
End Sub

Public Sub CargarModificacion(ID_Hoja As Long)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = " SELECT ID_HOJA_RUTA, FECHA, CANTIDADKILOMETROS, COD_PERSONAL,ESTADO"
    sql = sql & "  From HOJA_RUTA_CUERPO "
    sql = sql & "  Where ID_HOJA_RUTA=" & ID_Hoja
    
    grdDetalleHojaRuta.Enabled = True
  
    rs.Open sql, ConActiva, 0, 1
    If Not rs.EOF Then
        txtID_Hoja.Text = rs!ID_HOJA_RUTA
        txtFecha.Text = rs!Fecha
        ctlPersonal.Valor = rs!COD_PERSONAL
        ActualizarGrilla txtID_Hoja.Text
        If rs!ESTADO = 100 Then
            MsgBox "La hoja de ruta esta terminada " & vbCrLf & "No puede ser modificada", vbCritical
            grdDetalleHojaRuta.Enabled = False
        End If
        
    End If
    
  
End Sub

Public Sub UpdateHojaRuta()
    If Not IsNull(ctlPersonal.Valor) Then
        ExecutarSql " UPDATE HOJA_RUTA_CUERPO SET COD_PERSONAL = " & ctlPersonal.Valor & " Where ID_HOJA_RUTA = " & txtID_Hoja.Text
       AcualizarRequerimiento txtID_Hoja.Text
    Else
        MsgBox "El responsable esta en nulo"
    End If
End Sub


Public Sub AcualizarRequerimiento(HojaRuta As Long)

Dim sql As String
Dim requerimientosFiltro As String
Dim rs As New ADODB.Recordset


        
        sql = " Update Requerimiento"
            sql = sql & "  Set COD_HOJA_RUTA_TERMINADO = 0"
            sql = sql & "  WHERE  COD_HOJA_RUTA_TERMINADO = " & HojaRuta
            ExecutarSql sql
        
        
        sql = " SELECT     COD_HOJA_RUTA, COD_REQUERIMIENTO, ORDEN, KILOMETROS"
        sql = sql & "  From HOJA_RUTA_DETALLE"
         sql = sql & " Where COD_HOJA_RUTA = " & HojaRuta
        
        
       rs.Open sql, ConActiva, 0, 1
        
        Do While Not rs.EOF
             requerimientosFiltro = requerimientosFiltro & "," & Trim(rs!COD_REQUERIMIENTO)
            rs.MoveNext
        Loop

            sql = " Update Requerimiento"
            sql = sql & "  Set COD_HOJA_RUTA_TERMINADO = " & HojaRuta
            sql = sql & "  WHERE     IDREQUERIMIENTO IN (" & Mid(requerimientosFiltro, 2) & ")"
            ExecutarSql sql

End Sub

