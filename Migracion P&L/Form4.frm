VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form4"
   ScaleHeight     =   8460
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSinImagen 
      Caption         =   "sin Imagen"
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   900
      Width           =   1635
   End
   Begin VB.TextBox txtCaja 
      Height          =   435
      Left            =   4020
      TabIndex        =   8
      Top             =   780
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   435
      Left            =   9120
      TabIndex        =   7
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "excel"
      Height          =   435
      Left            =   9360
      TabIndex        =   6
      Top             =   300
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "buscar"
      Height          =   375
      Left            =   6300
      TabIndex        =   4
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox txtfechaCarga 
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtCliente 
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5835
      Left            =   660
      TabIndex        =   0
      Top             =   2340
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10292
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
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
            LCID            =   3082
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
            LCID            =   3082
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
      Caption         =   "Caja"
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha desde carga "
      Height          =   315
      Left            =   2220
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label cliente 
      Caption         =   "Cliente"
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CopiarDatosGrilla DataGrid1
End Sub

Private Sub Command2_Click()



Dim rs As New ADODB.Recordset
    Dim sql As String
   
    Dim conestado As New ADODB.Connection



Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"



sql = " SELECT   apellido, nombre, fecha, hora, minuto, idReferencia, idLoteReferencia, nombreCliente, clienteAsp_id, usuario_id, codigoContenedor, codigoElemento, fechaHora,"
sql = sql & "    Descripcion , numero1, numero2, texto1, texto2 , fecha1, fecha2, pathLegajo "
sql = sql & " From vista_referencia_usuario"
sql = sql & " WHERE (nombreCliente LIKE '%" & txtCliente.Text & "%') "
Rem sql = sql & " AND (fechaHora > CONVERT(DATETIME, '2016-03-30 00:00:00', 102))"
If txtfechaCarga.Text <> "" Then
    sql = sql & " AND fechaHora > '" & txtfechaCarga & "'"
End If
If txtCaja.Text <> "" Then
     sql = sql & " AND codigoContenedor = " & txtCaja.Text
End If

If chkSinImagen.Value = 1 Then
  sql = sql & "  AND (pathLegajo IS NULL)"
End If

sql = sql & " ORDER BY persona_id, fechaHora"

rs.CursorLocation = adUseClient
rs.Open sql, strConAsp150
Set DataGrid1.DataSource = rs.DataSource
    DataGrid1.Refresh

End Sub

Private Sub Command3_Click()
Dim sql As String

Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

Dim conBasa As New ADODB.Connection
Dim RSBASA As New ADODB.Recordset
Dim RSASP As New ADODB.Recordset

Dim ETIQUETA As String

conBasa.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"

sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.ID AS IDIMAGEN, DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA,"
sql = sql & vbCrLf & " DOCUMENTOS_DIGITALES.LOTE_ASP "
sql = sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
sql = sql & vbCrLf & " DOCUMENTOS_DIGITALES ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
sql = sql & vbCrLf & "  WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES IN (229, 5939)"
RSBASA.CursorLocation = adUseClient

RSBASA.Open sql, "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"


Do While Not RSBASA.EOF

ETIQUETA = ""
If Len(RSBASA!FK_LEGAJO_ETIQUETA) = 12 Then
    ETIQUETA = RSBASA!FK_LEGAJO_ETIQUETA
End If

If Len(RSBASA!FK_LEGAJO_ETIQUETA) = 13 Then
    ETIQUETA = Mid(RSBASA!FK_LEGAJO_ETIQUETA, 1, 12)
End If


If ETIQUETA <> "" Then

        sql = " SELECT        elementos_1.codigo AS CAJA, elementos.codigo AS ETIQUETA12, lotereferencia.codigo AS LOTE, referencia.pathLegajo, referencia.texto2, referencia.texto1,"
        sql = sql & vbCrLf & " referencia.numero2 , referencia.numero1"
        sql = sql & vbCrLf & " FROM            elementos INNER JOIN"
        sql = sql & vbCrLf & " referencia ON elementos.id = referencia.elemento_id INNER JOIN"
        sql = sql & vbCrLf & " lotereferencia ON referencia.lote_referencia_id = lotereferencia.id INNER JOIN"
        sql = sql & vbCrLf & " elementos AS elementos_1 ON elementos.contenedor_id = elementos_1.id"
        sql = sql & vbCrLf & " WHERE elementos.codigo = '" & ETIQUETA & "'"
        Set RSASP = New ADODB.Recordset
        
        RSASP.Open sql, strConAsp150
        
        If Not RSASP.EOF Then
        
        sql = " UPDATE DOCUMENTOS_DIGITALES Set LOTE_ASP = " & RSASP!LOTE
        sql = sql & " Where ID =" & RSBASA!IDIMAGEN
        conBasa.Execute sql
        
        End If
        
        

End If



    RSBASA.MoveNext
Loop

End Sub
