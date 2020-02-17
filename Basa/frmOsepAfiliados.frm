VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOsepAfiliados 
   Caption         =   "Osep Afiliados"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   9525
   Begin VB.TextBox txtDescripcion 
      Height          =   1635
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmOsepAfiliados.frx":0000
      Top             =   3780
      Width           =   7995
   End
   Begin VB.CommandButton cmdExpander 
      Caption         =   "Ver"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtapellido 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtDocumento 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtAfiliado 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   180
      Width           =   2415
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdOsepAfiliados 
      Height          =   1635
      Left            =   180
      TabIndex        =   0
      Top             =   1980
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   2884
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
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
   Begin VB.Label Label4 
      Caption         =   "Descripción :"
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   660
      TabIndex        =   7
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Documento"
      Height          =   255
      Left            =   660
      TabIndex        =   5
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Nº de Afiliado"
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   240
      Width           =   1635
   End
End
Attribute VB_Name = "frmOsepAfiliados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConOsep As ADODB.Connection
Private Sub cmdBuscar_Click()


        MousePointer = 11
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        Dim Sql As String
        rs.CursorLocation = adUseClient
       
        
        Sql = "    SELECT TIPO_DOC_AFI, NUMERO_AFI, VINCULO,"
        Sql = Sql & vbCrLf & " APELLIDO_NOMBRE, TIPO_DOC, DOCUMENTO, SITUACION,"
        Sql = Sql & vbCrLf & " DESCRIPCION"
        Sql = Sql & vbCrLf & " From OSEPAFILI"
        Sql = Sql & vbCrLf & " Where "
        If txtAfiliado.Text <> "" Then
            Sql = Sql & vbCrLf & " NUMERO_AFI = " & txtAfiliado.Text
        End If
        If txtDocumento.Text <> "" Then
            Sql = Sql & vbCrLf & " DOCUMENTO = " & txtDocumento.Text
        End If
        
        If txtapellido.Text <> "" And txtNombre.Text = "" Then
            Sql = Sql & vbCrLf & " APELLIDO_NOMBRE Like '%" & txtapellido.Text & "%'"
        End If
                
        If txtapellido.Text <> "" And txtNombre.Text <> "" Then
            Sql = Sql & vbCrLf & " APELLIDO_NOMBRE Like '%" & txtapellido.Text & "%' and APELLIDO_NOMBRE Like '%" & txtNombre.Text & "%'"
        End If
        
        rs.Open Sql, ConOsep
        If Not rs.EOF Then
            DATOSGRILLA grdOsepAfiliados, rs
             grdOsepAfiliados.Columns(6).WrapText = True
        Else
            MsgBox "No se encontraron registros"
        End If
       
        
        MousePointer = 0

End Sub

Private Sub cmdExpander_Click()
    txtDescripcion.Text = grdOsepAfiliados.Columns(7).Text
End Sub

Private Sub Form_Load()
    Set ConOsep = New ADODB.Connection
   
   ConOsep.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ClienteOsep & ";Persist Security Info=False"
   
End Sub
