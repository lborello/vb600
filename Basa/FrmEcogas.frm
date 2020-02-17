VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmEcogas 
   Caption         =   "Ecogas"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9000
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Datos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   60
      TabIndex        =   7
      Top             =   1980
      Width           =   8775
      Begin VB.TextBox txtBarrio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6060
         TabIndex        =   13
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtCalle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   10
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
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
         Left            =   7560
         TabIndex        =   8
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Barrio:"
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
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Width           =   795
      End
      Begin VB.Label CALLE 
         Caption         =   "Calle:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Numero:"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid grdCalle 
      Height          =   3195
      Left            =   60
      TabIndex        =   6
      Top             =   3000
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   5636
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
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
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   1755
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdActualizacion 
         Caption         =   "Actualizacion"
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
         Left            =   3540
         TabIndex        =   5
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtPaso 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         TabIndex        =   4
         Text            =   "C:\Actualizacion\"
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtFechaActualizacion 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Paso del Archivo"
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
         Left            =   180
         TabIndex        =   3
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de actualizacion:"
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
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiarGrilla 
         Caption         =   "Copiar Grilla"
      End
   End
End
Attribute VB_Name = "FrmEcogas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualizacion_Click()
    Dim sSQL As String
    Dim rs As New ADODB.Recordset
        MousePointer = 11
         Set rs = New ADODB.Recordset
            
         sSQL = " SELECT ID_CLIENTE_LEGAJO,COD_ESTADO, FECHA, COD_REMITO,CLIENTE_LEGAJO,"
         sSQL = sSQL & vbCrLf & " CONCAT(CONCAT(CONCAT(CONCAT(CONCAT(CALLE,CONCAT(' Nº:', NRO)),' PISO:'),PISO),' DTO:'),DEPTO) AS NOMBRE,"
         sSQL = sSQL & vbCrLf & " CONCAT(CONCAT(CONCAT('Bº:', ECOGASFINAL.BARRIO),'   LOC.:'), ECOGASFINAL.LOCALIDAD) AS DESCRIPCION"
         sSQL = sSQL & vbCrLf & " From LEGAJOS, ECOGASFINAL"
         sSQL = sSQL & vbCrLf & " Where LEGAJOS.ID_LEGAJO_ECOGAS = ECOGASFINAL.ID_EXPEDIENTE"
         sSQL = sSQL & vbCrLf & " AND (LEGAJOS.COD_CLIENTE = 4)  AND  (NOT (LEGAJOS.CLIENTE_LEGAJO IS NULL)) "
         Rem Ssql = Ssql & vbCrLf & " AND  ID_CLIENTE_LEGAJO IN (146845,146845,146845,146840,146839,146833,135699,132734,132733,128536,128120,100999,92352,83761,75408,1) "
         sSQL = sSQL & vbCrLf & "  AND (FECHA_ACTUALIZACION > " & FechaServerTipo(txtFechaActualizacion.Text)
         sSQL = sSQL & vbCrLf & "  ORDER BY ID_CLIENTE_LEGAJO "
         
         rs.Open sSQL, ConActiva, 0, 1
         rs.Save txtPaso & "Legajos.RST"
'         Dim sql As String
'         Dim con As New ADODB.Connection
'         con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Basa cliente\ClienteBasa.mdb;Jet OLEDB:Database Password=1742"
'
'         Do While Not RS.EOF
'         If Not IsNull(RS!COD_REMITO) Then
'         sql = "INSERT INTO  LEGAJOS (ID_CLIENTE_LEGAJO, CLIENTE_LEGAJO, NOMBRE,DESCRIPCION,COD_ESTADO,FECHA,COD_REMITO ) "
'        sql = sql & vbCrLf & " VALUES (" & RS!ID_CLIENTE_LEGAJO & ",'" & RS!CLIENTE_LEGAJO & "','" & RS!NOMBRE & "','" & RS!DESCRIPCION & "','" & RS!COD_ESTADO & "'," & "#" & Format(RS!FECHA, "dd/mm/yyyy") & "#" & "," & RS!COD_REMITO & ")"
'        con.Execute sql
'        Else
'         sql = "INSERT INTO  LEGAJOS (ID_CLIENTE_LEGAJO, CLIENTE_LEGAJO, NOMBRE,DESCRIPCION,COD_ESTADO ) "
'        sql = sql & vbCrLf & " VALUES (" & RS!ID_CLIENTE_LEGAJO & ",'" & RS!CLIENTE_LEGAJO & "','" & Replace(RS!NOMBRE, "|", " ") & "','" & RS!DESCRIPCION & "'," & RS!COD_ESTADO & ")"
'        con.Execute sql
'        End If
'        RS.MoveNext
'         Loop
         
         Set rs = New ADODB.Recordset
         sSQL = " SELECT ID_CLIENTE_LEGAJO, COD_ESTADO, COD_REMITO, FECHA  From LEGAJOS Where (COD_CLIENTE = 04) And (COD_ESTADO = 3) "
         rs.Open sSQL, ConActiva, 0, 1
         rs.Save txtPaso & "EstadoLegajos.RST"
         
'         Set RS = New ADODB.Recordset
'         Ssql = " SELECT INDICE, DESCRIPCION, FECHA, NUMERO, LETRA,EXPEDIENTE, APELLIDO_NOMBRE, MASK_EXPEDIENTE,MASK_LETRA, TOOLTIPFECHA, TOOLTIPEXPEDIENTE,TOOLTIPAPELLIDO_NOMBRE, TOOLTIPNUMERO,TOOLTIPLETRA , FECHA_MODIFICACION From INDICES Where (COD_CLIENTE = 04) ORDER BY INDICE"
'         RS.Open Ssql, strConBasa , 0 ,1
'         RS.Save "C:\Basa Cliente\INDICES.RST"
'
'         Set RS = New ADODB.Recordset
'         Ssql = " SELECT ESTADO,ID_UNITER,COD_CLIENTE, NRO_CAJA, ITEM, INDICE, DESCRIPCION, LETRA_DESDE, LETRA_HASTA, FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA, EXPEDIENTE,APELLIDO_NOMBRE, FECHA_MODIFICACION,USUARIO_MODIFICACION , BORRADO From REFERENCIAS Where (COD_CLIENTE = 04)"
'         RS.Open Ssql, strConBasa , 0 ,1
'         RS.Save "C:\Basa Cliente\REFERENCIAS.RST"
'
'         Set RS = New ADODB.Recordset
'         Ssql = " SELECT ESTADO, NRO_CAJA, NRO_REMITO, SECTOR, SOLICITANTE , FECHARECEPCION From CONTENEDOR, REQUERIMIENTO Where CONTENEDOR.NRO_REMITO = REQUERIMIENTO.IDREMITO (+) AND (CONTENEDOR.COD_CLIENTE = 04) ORDER BY CONTENEDOR.NRO_CAJA DESC"
'         RS.Open Ssql, strConBasa , 0 ,1
'         RS.Save "C:\Basa Cliente\ESTADOCAJAS.RST"
MousePointer = 0
     
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()

 MousePointer = 11
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim Sql As String
        rs.CursorLocation = adUseClient
        Sql = "   SELECT LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.NRO_CAJA, "
        Sql = Sql & vbCrLf & " ECOGASFINAL.NROEXPEDIENTE, ECOGASFINAL.CALLE, "
        Sql = Sql & vbCrLf & " ECOGASFINAL.NRO, ECOGASFINAL.PISO, "
        Sql = Sql & vbCrLf & " ECOGASFINAL.DEPTO, ECOGASFINAL.BARRIO, "
        Sql = Sql & vbCrLf & " LEGAJOS.Cod_Estado , ECOGASFINAL.LOCALIDAD "
        Sql = Sql & vbCrLf & " From ECOGASFINAL, LEGAJOS "
        Sql = Sql & vbCrLf & " Where ECOGASFINAL.ID_EXPEDIENTE = LEGAJOS.ID_LEGAJO_ECOGAS (+)"
        If Trim(txtCalle.Text) <> "" Then
            Sql = Sql & vbCrLf & " AND ECOGASFINAL.CALLE LIKE '%" & Trim(UCase(txtCalle.Text)) & "%'"
        End If
        If Trim(txtNumero.Text) <> "" Then
            Sql = Sql & vbCrLf & " AND ECOGASFINAL.NRO LIKE '%" & UCase(Trim(txtNumero.Text)) & "%'"
        End If
        If Trim(txtBarrio.Text) <> "" Then
            Sql = Sql & vbCrLf & " AND ECOGASFINAL.BARRIO LIKE '%" & Trim(UCase(txtBarrio.Text)) & "%'"
        End If
        rs.Open Sql, ConActiva, 0, 1
        If Not rs.EOF Then
            DATOSGRILLA grdCalle, rs
        End If
        MousePointer = 0
End Sub

Private Sub grdCalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnu
    End If
End Sub

Private Sub mnuCopiarGrilla_Click()
    CopiarDatosGrilla grdCalle
End Sub
