VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C30A4F2E-16E3-4694-9920-512C55E5C51A}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmControlesVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
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
   ScaleHeight     =   5835
   ScaleWidth      =   7845
   Begin Controles.cltGenerico ctlClientes 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   420
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Control"
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
      Left            =   6540
      TabIndex        =   4
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton cmdCopiarExcel 
      Caption         =   "Copiar Excel"
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
      Left            =   6060
      TabIndex        =   3
      Top             =   5340
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid grdControl 
      Height          =   4395
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7752
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
   Begin VB.ComboBox cboTipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmControlesVarios.frx":0000
      Left            =   1020
      List            =   "frmControlesVarios.frx":0010
      TabIndex        =   0
      Top             =   60
      Width           =   5355
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmControlesVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdControl
End Sub

Private Sub Command1_Click()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        rs.CursorLocation = adUseClient
        MousePointer = 11
       Select Case cboTipo.ListIndex
        Case 0
                Sql = "  SELECT ESTANTERIA , ESTADO, COD_CLIENTE, NRO_CAJA,"
                Sql = Sql & vbCrLf & "  F_MODIFICACION , NRO_REMITO, UB_PROVISORIA"
                Sql = Sql & vbCrLf & "  From CONTENEDOR"
                Sql = Sql & vbCrLf & "  WHERE (ESTANTERIA BETWEEN 109 AND 200) AND"
                Sql = Sql & vbCrLf & "  (ESTADO IN (2)) AND (COD_CLIENTE <> 39) AND"
                Sql = Sql & vbCrLf & "  (UB_PROVISORIA IS NULL)"
                Sql = Sql & vbCrLf & "  ORDER BY COD_CLIENTE, F_MODIFICACION"
                
        
        Case 1
                Sql = " SELECT ESTANTERIA, ESTADO, COD_CLIENTE, NRO_CAJA, F_MODIFICACION , NRO_REMITO"
                Sql = Sql & vbCrLf & "  From CONTENEDOR"
                Sql = Sql & vbCrLf & "  WHERE (NOT (ESTANTERIA BETWEEN 109 AND 200)) AND "
                Sql = Sql & vbCrLf & "  (ESTADO IN (5,4 )) AND (COD_CLIENTE <> 39) AND "
                Sql = Sql & vbCrLf & "  (UB_PROVISORIA IS NULL)ORDER BY COD_CLIENTE, F_MODIFICACION "
       Case 2
       
        If IsNull(ctlClientes.Valor) Then
            MsgBox "Ingrese el cliente"
            Exit Sub
        End If
        
            Sql = " SELECT COD_INDICE, CLIENTE_LEGAJO, COUNT(*) AS CANT"
            Sql = Sql & vbCrLf & " From LEGAJOS"
            Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & ctlClientes.Valor
            Sql = Sql & vbCrLf & " GROUP BY COD_INDICE, CLIENTE_LEGAJO"
            Sql = Sql & vbCrLf & " HAVING (COUNT(*) > 1)"
                
         Case 3
                Sql = " SELECT     CONTENEDOR.ESTANTERIA, CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CONTENEDOR.F_MODIFICACION,"
                Sql = Sql & vbCrLf & " CONTENEDOR.NRO_REMITO , CLIENTEUSUARIO.APELLIDO_NOMBRE, REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.IDREMITO"
                Sql = Sql & vbCrLf & " FROM         CLIENTEUSUARIO INNER JOIN"
                Sql = Sql & vbCrLf & " REQUERIMIENTO ON CLIENTEUSUARIO.ID_CLIENTEUSUARIO = REQUERIMIENTO.COD_USUARIO_CLIENTE RIGHT OUTER JOIN"
                Sql = Sql & vbCrLf & " CONTENEDOR INNER JOIN"
                Sql = Sql & vbCrLf & " REQUELIBOSCAJAS ON CONTENEDOR.NRO_CAJA = REQUELIBOSCAJAS.CAJASLIBROS ON"
                Sql = Sql & vbCrLf & " REQUERIMIENTO.IDREQUERIMIENTO = REQUELIBOSCAJAS.IDREQUERIMIENTOS AND"
                Sql = Sql & vbCrLf & " REQUERIMIENTO.id_cliente = CONTENEDOR.COD_CLIENTE"
                Sql = Sql & vbCrLf & " WHERE     (NOT (CONTENEDOR.ESTANTERIA BETWEEN 109 AND 200)) AND (CONTENEDOR.ESTADO IN (5, 4)) AND (CONTENEDOR.COD_CLIENTE <> 39) AND"
                Sql = Sql & vbCrLf & " (CONTENEDOR.UB_PROVISORIA IS NULL) AND (REQUERIMIENTO.IDTIPOREQUERIMIENTO = 7)"
                Sql = Sql & vbCrLf & " ORDER BY CONTENEDOR.COD_CLIENTE, CONTENEDOR.F_MODIFICACION"
         
        End Select
        
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdControl, rs
        MousePointer = 0
End Sub

Private Sub Form_Load()
    ctlClientes.TipoControl = Cliente
End Sub
