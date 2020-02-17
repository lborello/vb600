VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indice"
   ClientHeight    =   9120
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11385
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
   ScaleHeight     =   9120
   ScaleWidth      =   11385
   Begin VB.CommandButton cmdBuscarLegajos 
      Caption         =   "Buscar Legajos"
      Height          =   315
      Left            =   5100
      TabIndex        =   2
      Top             =   60
      Width           =   1575
   End
   Begin MSComctlLib.TreeView trvIndices 
      Height          =   8475
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   14949
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
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
   Begin VB.TextBox txtBuscarIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   0
      Top             =   60
      Width           =   2715
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":0279
            Key             =   "Documento"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":04FF
            Key             =   "Sector"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":08BC
            Key             =   "Documentos"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":0B3D
            Key             =   "Legajo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":1417
            Key             =   "Sucursal"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndice.frx":17EA
            Key             =   "Casa"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Busqueda descripción:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuIndice 
      Caption         =   "mnuIndice"
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
      End
      Begin VB.Menu mnuBuscarLegajos 
         Caption         =   "Buscar legajos"
      End
   End
End
Attribute VB_Name = "frmIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ControlAsignacion As Control
Public COD_CLIENTE As Integer


Private Sub cltIndice1_Click()

End Sub

Private Sub ctlIndice_DblClick()
   

End Sub

Private Sub cmdBuscarLegajos_Click()
    BuscarTipoIndice "0", "Legajo", True
End Sub

Private Sub Form_Load()
frmIndice.Top = 0
frmIndice.Left = 0
End Sub

Private Sub mnuBuscar_Click()
Rem ctlIndice.BuscarIndice InputBox("Texto a Buscar"), True

End Sub

Private Sub mnuBuscarLegajos_Click()
BuscarTipoIndice Mid(trvIndices.SelectedItem.Key, 2), "Legajo", True
End Sub

Public Sub COntroles(contr As Control)

End Sub

Private Sub trvIndices_DblClick()
   On Error GoTo salir
    itemSelecionado = Item_Selecionado
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Sql = " SELECT DESCRIPCION, ID_CODIGO_DOCUMENTO, ID "
    Sql = Sql & " From INDICES"
    Sql = Sql & " WHERE COD_CLIENTE = " & COD_CLIENTE
    Sql = Sql & " AND INDICE =  '" & Item_Selecionado & "'"
    rs.Open Sql, ConActiva, 0, 1
    frmAgregarDocumentos.NºDOCUMENTO = rs!ID_CODIGO_DOCUMENTO
    Rem ControlAsignacion.Text = rs!ID_CODIGO_DOCUMENTO
    NRO_DOCUMENTO = rs!ID_CODIGO_DOCUMENTO
    
    Unload Me
salir:
End Sub

Private Sub trvIndices_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuIndice
   End If
End Sub

Private Sub txtBuscarIndice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscarIndice txtBuscarIndice, True
End If
End Sub
Public Function Item_Selecionado() As String
Dim Nodo_1 As Node
    Dim i As Integer
       With trvIndices.Nodes
           For i = 1 To .Count
              If .Item(i).Selected Then
                  Item_Selecionado = Mid(.Item(i).Key, 2)
                  Exit Function
              End If
           Next
        End With
  End Function
Public Function Index_Selecionado() As Integer
Dim Nodo_1 As Node
    Dim i As Integer
       With trvIndices.Nodes
           For i = 1 To .Count
              If .Item(i).Selected Then
                 Index_Selecionado = i
                  Exit Function
              End If
           Next
        End With
  End Function
  Public Function Descripcion() As String
Dim Nodo_1 As Node
    Dim i As Integer
       With trvIndices.Nodes
           For i = 1 To .Count
              If .Item(i).Selected Then
                 Descripcion = .Item(i).Text
                  Exit Function
              End If
           Next
        End With
  End Function

  
'Public Function Indice() As String
'Dim i As Integer
'Dim ItemSel As String
'ItemSel = Item_Selecionado
' For i = 1 To Len(ItemSel)
'  If Mid(ItemSel, i, 1) = "-" Then
'    Indice = Trim(Mid(ItemSel, 1, i - 1))
'  End If
'
'
'  Next
'  End Function
    
  
  Public Function Cliente() As Integer
    Cliente = COD_CLIENTE
  End Function
Public Function Actualizar(COD_CLIENTE As Integer, Filtro As TipoIndice, ExpanderIndex As Integer, Optional Filtro_Indice As String)
    Dim Indice0 As String
    Dim KeyTreeView1 As String
    Dim Indice1 As String
    Dim Descripcion As String
    Dim nodX As Node
    Dim rsIndices As New ADODB.Recordset
    Dim Sql As String
    Dim Tipo_Indice As String
   COD_CLIENTE = COD_CLIENTE
  If Filtro_Indice = "" Then
     Sql = " SELECT * From INDICES  Where COD_CLIENTE =" & COD_CLIENTE
   Else
    Sql = " SELECT * From INDICES  Where COD_CLIENTE =" & COD_CLIENTE
    Sql = Sql & " AND  (INDICE LIKE '" & Filtro_Indice & "%')"
   End If

    Select Case Filtro
    Case TipoIndice.Sector
        Sql = Sql & " AND TIPO_INDICE ='Sector' "
    End Select

    Sql = Sql & " ORDER BY INDICE"
      rsIndices.Open Sql, ConActiva, 0, 1

        trvIndices.Nodes.Clear
        Set nodX = trvIndices.Nodes.Add(, , "RAIZ", "TODAS LAS CATEGORIAS", "Casa") ' Root
        trvIndices.Nodes.Item("RAIZ").Tag = "TODOS"
        Do While Not rsIndices.EOF
        Tipo_Indice = Trim(rsIndices!Tipo_Indice)

            If ExisItem("R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)) Then
                KeyTreeView1 = "R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)
                Descripcion = rsIndices!Indice & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!Descripcion)
                Set nodX = trvIndices.Nodes.Add(KeyTreeView1, tvwChild, "R" & rsIndices!Indice, Descripcion, Tipo_Indice, Tipo_Indice)
                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
            Else
                Descripcion = rsIndices!Indice & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!Descripcion)
                Set nodX = trvIndices.Nodes.Add(, , "R" & rsIndices!Indice, Descripcion, Tipo_Indice, Tipo_Indice) ' Root
                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
            End If
            rsIndices.MoveNext
        Loop
        EXPANDER ExpanderIndex
End Function
Public Function EXPANDER(IndiceNumerico As Integer)
Dim i As Integer
Dim indexs As Integer
indexs = IndiceNumerico
On Error Resume Next
For i = 1 To 10
   indexs = trvIndices.Nodes.Item(indexs).Parent.Index
   trvIndices.Nodes.Item(indexs).Expanded = True
Next
trvIndices.Nodes.Item(IndiceNumerico).Expanded = True
trvIndices.Refresh
End Function
Public Function ExisItem(DATO As String) As Boolean
    Dim s As String
    On Error GoTo ErrorHandler
        ExisItem = True
        
        s = trvIndices.Nodes.Item(DATO)
    Exit Function
ErrorHandler:
    ExisItem = False
End Function
Private Function PonerImagen(Doc As Variant) As Integer
PonerImagen = 0
If Not IsNull(Doc) Then
    Select Case Trim(Doc)
    Case "Sector"
        PonerImagen = 3
    Case "Documento"
        PonerImagen = 2
    Case "Documentos"
        PonerImagen = 4
    Case "Legajo"
        PonerImagen = 5
        
    End Select
End If
End Function
'Public Sub MarcarIndiceFrase(Dato As String, Optional EXPANDER As Boolean)
'    Dim i  As Integer
'    Dim a As Integer
'    Dim B As Integer
'    Dim Indice As String
'
'
'        For i = 1 To trvIndices.Nodes.Count
'            trvIndices.Nodes.Item(i).BackColor = &H80000005
'            If Dato = "" Or Dato = " " Then
'            Else
'                B = InStr(UCase(trvIndices.Nodes.Item(i).Text), "-")
'                If UCase(trvIndices.Nodes.Item(i).Text) <> "TODAS LAS CATEGORIAS" Then
'                   If Mid(Dato, 1, 1) = "0" Then
'                       ' BUSCAR INDICE
'                       Indice = Mid(trvIndices.Nodes.Item(i).Text, 1, B - 2)
'                       If Indice = UCase(Dato) Then
'                            a = 1
'                       Else
'                            a = 0
'                       End If
'                   Else
'                        ' BUSCAR NOMBRE
'                        a = InStr(UCase(trvIndices.Nodes.Item(i).Text), UCase(Dato))
'                    End If
'                    If a = 0 Then
'                      If EXPANDER = True Then
'                        trvIndices.Nodes.Item(i).Expanded = False
'                      End If
'                    Else
''                        trvIndices.Nodes.Item(i).Expanded = True
''                        trvIndices.Nodes.Item(i).Selected = True
''                        trvIndices.Nodes.Item(i).BackColor = &H80000002
''                        trvIndices.Nodes.Item(i).Bold = True
'                    End If
'                End If
'            End If
'        Next
'End Sub
Public Sub BuscarIndice(DATO As String, Optional EXPANDER As Boolean)
    Dim i  As Integer
    Dim a As Integer
    Dim B As Integer
    Dim Indice As String
        Dim ImagenLegajo As Integer
        
        For i = 1 To trvIndices.Nodes.Count
            trvIndices.Nodes.Item(i).BackColor = &H80000005
            trvIndices.Nodes.Item(i).ForeColor = &H80000008
            trvIndices.Nodes.Item(i).Bold = False
            If Trim(DATO) = "" Then
                Exit Sub
            Else
                    a = InStr(UCase(trvIndices.Nodes.Item(i).Text), UCase(DATO))
                    If a <> 0 Then
                              If EXPANDER = True Then
                                    trvIndices.Nodes.Item(i).ForeColor = &H80000002
                                    trvIndices.Nodes.Item(i).Bold = True
                                    trvIndices.Nodes.Item(i).Selected = True
                                    trvIndices.Nodes.Item(i).Expanded = True
                              Else
                                    trvIndices.Nodes.Item(i).ForeColor = &H80000002
                                    trvIndices.Nodes.Item(i).Bold = True
                                    trvIndices.Nodes.Item(i).Selected = True
                                    trvIndices.Nodes.Item(i).Expanded = True
                              End If
                    End If
             End If
                    
        Next
        trvIndices.Refresh
End Sub
Public Sub BuscarTipoIndice(Indice As String, DATO As String, Optional EXPANDER As Boolean)
    Dim i  As Integer
    Dim a As Integer
    Dim B As Integer

        Dim ImagenLegajo As Integer
        
        For i = 1 To trvIndices.Nodes.Count
            trvIndices.Nodes.Item(i).BackColor = &H80000005
            trvIndices.Nodes.Item(i).ForeColor = &H80000008
            trvIndices.Nodes.Item(i).Bold = False
            If trvIndices.Nodes.Item(i).Image = DATO Then
               If Indice = Mid(trvIndices.Nodes.Item(i).Key, 2, Len(Indice)) Then
                  If EXPANDER = True Then
                        trvIndices.Nodes.Item(i).ForeColor = &H80000002
                        trvIndices.Nodes.Item(i).Bold = True
                        trvIndices.Nodes.Item(i).Selected = True
                        trvIndices.Nodes.Item(i).Expanded = True
                  Else
                        trvIndices.Nodes.Item(i).ForeColor = &H80000002
                        trvIndices.Nodes.Item(i).Bold = True
                        trvIndices.Nodes.Item(i).Selected = True
                        trvIndices.Nodes.Item(i).Expanded = True
                  End If
              End If
            End If
        Next
        trvIndices.Refresh
End Sub

