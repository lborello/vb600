VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada"
   ClientHeight    =   8700
   ClientLeft      =   3480
   ClientTop       =   1470
   ClientWidth     =   8865
   Icon            =   "frmEntrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   8865
   Begin TabDlg.SSTab SSTab1 
      Height          =   8475
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   14949
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmEntrada.frx":0152
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblError"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cantElem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cantLegajos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cantLibros"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cantCajas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "gr_Caja"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "gr_libros"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "gr_legajos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCancelar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAceptar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtResponsable"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdBorrarCaja"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdLibros"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdBorrarLegajos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmEntrada.frx":016E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Elemento"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "grdEntrada"
      Tab(1).Control(3)=   "txtElemento"
      Tab(1).Control(4)=   "cmdBuscar"
      Tab(1).Control(5)=   "txtLote"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmEntrada.frx":018A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdBorrarLegajos 
         Caption         =   "Borrar Legajos"
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
         Left            =   6360
         TabIndex        =   24
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton cmdLibros 
         Caption         =   "Borrar Libros"
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
         Left            =   3600
         TabIndex        =   23
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton cmdBorrarCaja 
         Caption         =   "Borrar Cajas"
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
         Left            =   600
         TabIndex        =   22
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
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
         Left            =   -71100
         TabIndex        =   20
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton cmdBuscar 
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
         Height          =   315
         Left            =   -68700
         TabIndex        =   18
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox txtElemento 
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
         Left            =   -73800
         TabIndex        =   17
         Top             =   540
         Width           =   1155
      End
      Begin MSDataGridLib.DataGrid grdEntrada 
         Height          =   7155
         Left            =   -74820
         TabIndex        =   16
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   12621
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin VB.TextBox txtResponsable 
         BackColor       =   &H00FFFFFF&
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
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1575
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
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   7920
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   7140
         TabIndex        =   1
         Top             =   7920
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid gr_legajos 
         Height          =   6015
         Left            =   5700
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColorSel    =   8454016
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gr_libros 
         Height          =   6015
         Left            =   2940
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gr_Caja 
         Height          =   6015
         Left            =   60
         TabIndex        =   5
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Lote"
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
         Left            =   -71820
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Elemento 
         Caption         =   "Elemento"
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
         Left            =   -74760
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de Elementos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Cajas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Libros:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Legajos:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5820
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label cantCajas 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label cantLibros 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4380
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label cantLegajos 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7140
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label cantElem 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   5520
         TabIndex        =   8
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblError 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   4920
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
Dim ConBasa As ADODB.Connection
Dim sSQL As String
Dim ssql1 As String
Dim fecha As String
On Error GoTo ErrorHandler
Dim ConEntrada As New ADODB.Connection

ConEntrada.Open strConBasa
ConEntrada.CursorLocation = adUseClient
   Dim rs As New ADODB.Recordset
   Dim lote As Long
   
   
  rs.Open " SELECT  MAX(LOTE) AS MAXLOTE From dbo.ENTRADA ", ConActiva, 0, 1
   lote = rs!MaxLote + 1
    
   
    sSQL = "insert into entrada ( cod_cliente, elemento, tipo, fecha, cod_personal, cod_estado, lOTE) values (  "
    fecha = SysDate
    
   
    
    'Cajas
        For i = 1 To gr_Caja.Rows - 1
             ssql1 = gr_Caja.TextMatrix(i, 1) & ", " & gr_Caja.TextMatrix(i, 2) & ", " & 0 & ", " & fecha & "," & MDIfrmInicio.StaInicio.Panels.Item(2).Text & " , 0," & lote & ") "
            ConEntrada.Execute sSQL + ssql1
        Next
    'Libros
        For i = 1 To gr_libros.Rows - 1
            ssql1 = gr_libros.TextMatrix(i, 1) & ", " & gr_libros.TextMatrix(i, 2) & ", " & 1 & ", " & fecha & ", " & MDIfrmInicio.StaInicio.Panels.Item(2).Text & ",0," & lote & ") "
            ConEntrada.Execute sSQL + ssql1
        Next
    'Legajos
        For i = 1 To gr_legajos.Rows - 1
            ssql1 = gr_legajos.TextMatrix(i, 1) & ", " & gr_legajos.TextMatrix(i, 2) & ", " & 3 & ", " & fecha & ", " & MDIfrmInicio.StaInicio.Panels.Item(2).Text & " , 0," & lote & ") "
            ConEntrada.Execute sSQL + ssql1
        Next
       


     
     MsgBox "Registros Almacenados en el lote " & lote, vbInformation
     
     Call limpiar
     Exit Sub
ErrorHandler:


    
   

    If Err <> 0 Then
        MsgBox Err.Source & "-->" & Err.Description, , "Error"
    End If

    
    
End Sub

Private Sub cmdBorrarCaja_Click()
gr_Caja.Clear
gr_Caja.Rows = 1
gr_Caja.ColWidth(0) = 100
gr_Caja.ColAlignment(1) = 4
gr_Caja.ColAlignment(2) = 4
gr_Caja.TextMatrix(0, 1) = "Cliente"
gr_Caja.TextMatrix(0, 2) = "Caja"


End Sub

Private Sub cmdBorrarLegajos_Click()
gr_legajos.Clear
gr_legajos.Rows = 1

gr_legajos.ColWidth(0) = 100
gr_legajos.ColAlignment(1) = 4
gr_legajos.ColAlignment(2) = 4
gr_legajos.TextMatrix(0, 1) = "Cliente"
gr_legajos.TextMatrix(0, 2) = "Legajo"
End Sub

Private Sub cmdBuscar_Click()

 Dim rs As New ADODB.Recordset
    Dim Sql As String
    
    rs.CursorLocation = adUseClient
        
        Sql = " SELECT     ID_ENTRADA, COD_CLIENTE, ELEMENTO, TIPO, FECHA, COD_PERSONAL, COD_ESTADO, FECHA_LIMPIADO, LOTE"
        Sql = Sql & "  From ENTRADA"
        Sql = Sql & "  Where "
        If txtElemento.Text <> "" Then
           Sql = Sql & "   Elemento = " & txtElemento.Text
        Else
            If txtLote.Text <> "" Then
                Sql = Sql & " LOTE = " & txtLote.Text
            End If
        End If
        
        
        
    Sql = Sql & "  order by COD_CLIENTE "
        
        
       rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
       
       Set grdEntrada.DataSource = rs.DataSource
       grdEntrada.Refresh




End Sub

Private Sub cmdCancelar_Click()
End

End Sub

Private Sub Command1_Click()
Dim i As Integer

For i = 1244 To 1335
     gr_libros.AddItem "" & vbTab & 115 & vbTab & i
Next
 
End Sub

Private Sub cmdLibros_Click()
gr_libros.Clear
gr_libros.Rows = 1
gr_libros.ColWidth(0) = 100
gr_libros.ColAlignment(1) = 4
gr_libros.ColAlignment(2) = 4
gr_libros.TextMatrix(0, 1) = "Cliente"
gr_libros.TextMatrix(0, 2) = "Libro"


End Sub

Private Sub Form_Load()
gr_Caja.ColWidth(0) = 100
gr_Caja.ColAlignment(1) = 4
gr_Caja.ColAlignment(2) = 4
gr_Caja.TextMatrix(0, 1) = "Cliente"
gr_Caja.TextMatrix(0, 2) = "Caja"

gr_libros.ColWidth(0) = 100
gr_libros.ColAlignment(1) = 4
gr_libros.ColAlignment(2) = 4
gr_libros.TextMatrix(0, 1) = "Cliente"
gr_libros.TextMatrix(0, 2) = "Libro"

gr_legajos.ColWidth(0) = 100
gr_legajos.ColAlignment(1) = 4
gr_legajos.ColAlignment(2) = 4
gr_legajos.TextMatrix(0, 1) = "Cliente"
gr_legajos.TextMatrix(0, 2) = "Legajo"


End Sub

Private Sub Text1_Change()

End Sub




Private Sub txtResponsable_KeyPress(KeyAscii As Integer)

Dim rs As ADODB.Recordset
Dim Etiqueta As Long
Dim Cliente As Integer
Dim Sql As String

On Error GoTo salir:
If KeyAscii = 13 Then


If (UCase(Mid(txtResponsable.Text, 1, 2)) = "L1" Or Len(txtResponsable.Text) = 11) And UCase(Mid(txtResponsable.Text, 1, 2)) <> "C5" Then
            For j = 1 To gr_legajos.Rows - 1
                If Not IsNumeric(Mid(txtResponsable.Text, 6, 6)) Then
                MsgBox "error repita la operacion"
                txtResponsable.Text = ""
                Exit Sub
                End If
                
            
                 If gr_legajos.TextMatrix(j, 1) = Mid(txtResponsable.Text, 3, 3) And gr_legajos.TextMatrix(j, 2) = Mid(txtResponsable.Text, 6, 6) Then
                        lblError.Caption = "Error ya se ingreso el Legajo  " & txtResponsable.Text
                        lblError.BackColor = &H8080FF
                        txtResponsable.Text = ""
                    Exit Sub
                 End If
            Next
           
           gr_legajos.AddItem ("" & vbTab & Mid(txtResponsable.Text, 3, 3) & vbTab & Mid(txtResponsable.Text, 6))
           cantLegajos.Caption = gr_legajos.Rows - 1
           KeyAscii = 0
     End If
     
     If (UCase(Mid(txtResponsable.Text, 1, 2)) = "L2" And Len(txtResponsable.Text) = 9) Or (UCase(Mid(txtResponsable.Text, 1, 2)) = "12" And Len(txtResponsable.Text) = 13 Or Len(txtResponsable.Text) = 12) Then
            
            
            If Len(txtResponsable.Text) = 9 Then
                    Etiqueta = Mid(txtResponsable.Text, 3)
                   
             End If
             
             If Len(txtResponsable.Text) = 13 Then
                Etiqueta = Mid(txtResponsable.Text, 3, 10)
             End If
             
             
             If Len(txtResponsable.Text) = 12 And Mid(txtResponsable.Text, 1, 2) = "12" Then
                Etiqueta = Mid(txtResponsable.Text, 3, 10)
             End If
             
             If Not IsNumeric(Etiqueta) Then
                MsgBox "error repita la operacion"
                txtResponsable.Text = ""
                Exit Sub
                End If

            Set rs = New ADODB.Recordset
            Sql = " SELECT     ID_CLIENTE_LEGAJO, COD_CLIENTE , COD_ESTADO "
            Sql = Sql & " From LEGAJOS  "
                Sql = Sql & "  Where ID_LEGAJO = " & Etiqueta
            rs.Open Sql, ConActiva, 0, 1
            If rs.EOF Then
                MsgBox "Error Legajo"
                Exit Sub
            Else
            
            If IsNull(rs!COD_CLIENTE) Then
                MsgBox "Error Legajo"
                Exit Sub
            
            Else
               Cliente = rs!COD_CLIENTE
               If rs!Cod_Estado <> 3 Then
                    txtResponsable.Text = ""
                    MsgBox "Error en Estado ", vbCritical
                    Exit Sub
               End If
               
            End If
            End If
            
            
            
            For j = 1 To gr_legajos.Rows - 1
                 If gr_legajos.TextMatrix(j, 1) = Cliente And gr_legajos.TextMatrix(j, 2) = rs!ID_CLIENTE_LEGAJO Then
                        lblError.Caption = "Error ya se ingreso el Legajo  " & rs!ID_CLIENTE_LEGAJO
                        lblError.BackColor = &H8080FF
                        txtResponsable.Text = ""
                    Exit Sub
                 End If
            Next
           
           gr_legajos.AddItem ("" & vbTab & Cliente & vbTab & rs!ID_CLIENTE_LEGAJO)
           cantLegajos.Caption = gr_legajos.Rows - 1
           KeyAscii = 0
     End If
     
     

If UCase(Mid(txtResponsable.Text, 1, 2)) = "C5" Or UCase(Mid(txtResponsable.Text, 1, 2)) = "11" Then
            Dim Caja As Long
       
Set rs = New ADODB.Recordset
Caja = 0

    If Len(txtResponsable.Text) = 13 Then
                Caja = Mid(txtResponsable.Text, 3, 10)
    End If
    
            Sql = " SELECT     FK_CLIENTE, NRO_CAJA, ID_CAJA"
            Sql = Sql & " From dbo.Cajas "
            Sql = Sql & " Where ID_CAJA = " & Caja
            
            rs.Open Sql, ConActiva, 0, 1
            
            If Not rs.EOF Then
                Caja = rs!NRO_CAJA
                Cliente = rs!FK_CLIENTE
            
            End If
            
            
            
            
            For j = 1 To gr_Caja.Rows - 1
                 If gr_Caja.TextMatrix(j, 1) = Cliente And gr_Caja.TextMatrix(j, 2) = Caja Then
                        lblError.Caption = "Error ya se ingreso la Caja " & Caja
                        lblError.BackColor = &H8080FF
                        txtResponsable.Text = ""
                    Exit Sub
                 End If
            Next
     
            gr_Caja.AddItem "" & vbTab & Cliente & vbTab & Caja
            cantCajas.Caption = gr_Caja.Rows - 1
            
            KeyAscii = 0
    End If
     
     If (UCase(Mid(txtResponsable.Text, 1, 2)) = "C1" And Len(txtResponsable.Text) = 10) Then
            For j = 1 To gr_Caja.Rows - 1
                 If gr_Caja.TextMatrix(j, 1) = Mid(txtResponsable.Text, 3, 3) And gr_Caja.TextMatrix(j, 2) = Mid(txtResponsable.Text, 6) Then
                        lblError.Caption = "Error ya se ingreso la Caja " & txtResponsable.Text
                        lblError.BackColor = &H8080FF
                        txtResponsable.Text = ""
                    Exit Sub
                 End If
            Next
     
            gr_Caja.AddItem "" & vbTab & Mid(txtResponsable.Text, 3, 3) & vbTab & Mid(txtResponsable.Text, 6)
            cantCajas.Caption = gr_Caja.Rows - 1
            
            KeyAscii = 0
     Else
        If (Len(txtResponsable.Text) >= 18 And Len(txtResponsable.Text) <= 20) Then
            For j = 1 To gr_Caja.Rows - 1
                 If gr_Caja.TextMatrix(j, 1) = Mid(txtResponsable.Text, Len(txtResponsable.Text) - 9, 4) And gr_Caja.TextMatrix(j, 2) = Mid(txtResponsable.Text, Len(txtResponsable.Text) - 5) Then
                        lblError.Caption = "Error ya se ingreso la Caja  " & txtResponsable.Text
                        lblError.BackColor = &H8080FF
                        txtResponsable.Text = ""
                    Exit Sub
                 End If
            Next
            gr_Caja.AddItem "" & vbTab & Mid(txtResponsable.Text, Len(txtResponsable.Text) - 9, 4) & vbTab & Mid(txtResponsable.Text, Len(txtResponsable.Text) - 5)
            cantCajas.Caption = gr_Caja.Rows - 1
            KeyAscii = 0
        End If
     End If
     If Len(txtResponsable.Text) = 15 Then
            For j = 1 To gr_libros.Rows - 1
                 If gr_libros.TextMatrix(j, 1) = Mid(txtResponsable.Text, 11, 5) And gr_libros.TextMatrix(j, 2) = Mid(txtResponsable.Text, 6, 5) Then
                        lblError.Caption = "Error ya se ingreso el libro  " & txtResponsable.Text
                        lblError.BackColor = &H8080FF
                        txtResponsable.Text = ""
                    Exit Sub
                 End If
            Next
            gr_libros.AddItem "" & vbTab & Mid(txtResponsable.Text, 11, 5) & vbTab & Mid(txtResponsable.Text, 6, 5)
            cantLibros.Caption = gr_libros.Rows - 1
            
     End If
     cantElem.Caption = CInt(cantLibros.Caption) + CInt(cantCajas.Caption) + CInt(cantLegajos.Caption)
 lblError.BackColor = &H8000000F
 lblError.Caption = ""
txtResponsable.Text = ""
 
 

 
 
 End If
 
 
 Exit Sub
 
 
salir:
 MsgBox "Errror " & Err.Description
 
 
 

End Sub
Public Sub limpiar()
    gr_libros.Clear
    gr_Caja.Clear
    gr_legajos.Clear
    gr_libros.Rows = 1
    gr_Caja.Rows = 1
    gr_legajos.Rows = 1
    gr_Caja.ColWidth(0) = 100
    gr_Caja.TextMatrix(0, 1) = "Cliente"
    gr_Caja.TextMatrix(0, 2) = "Caja"

    gr_libros.ColWidth(0) = 100
    gr_libros.TextMatrix(0, 1) = "Cliente"
    gr_libros.TextMatrix(0, 2) = "Libro"

    gr_legajos.ColWidth(0) = 100
    gr_legajos.TextMatrix(0, 1) = "Cliente"
    gr_legajos.TextMatrix(0, 2) = "Legajo"
    cantCajas.Caption = 0
    cantLegajos.Caption = 0
    cantLibros.Caption = 0
    cantElem.Caption = 0
End Sub
