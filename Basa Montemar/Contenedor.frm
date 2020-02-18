VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContenedor 
   Caption         =   "COntenedor"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   13350
   Begin TabDlg.SSTab SSTab1 
      Height          =   8355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   14737
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Contenedor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtEstado1"
      Tab(0).Control(1)=   "txtEstanteriaHasta2"
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(3)=   "txtEstanteriaDesde1"
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(5)=   "DataGrid1"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Label1"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Contenedor.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cboTipoCaja"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdCrear"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtVerticalHasta"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtHorinzontalHasta"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtEstanteriaHasta"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtEstado"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtVerticalDesde"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtEstanteriaDesde"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtHorinzontalDesde"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command3"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Contenedor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   25
         Top             =   4260
         Width           =   1335
      End
      Begin VB.TextBox txtEstado1 
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
         Left            =   -67800
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEstanteriaHasta2 
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
         Left            =   -71220
         TabIndex        =   22
         Text            =   "0"
         Top             =   480
         Width           =   1875
      End
      Begin VB.CommandButton Command2 
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
         Height          =   315
         Left            =   -63720
         TabIndex        =   21
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox txtHorinzontalDesde 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1980
         TabIndex        =   20
         Text            =   "22"
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox txtEstanteriaDesde1 
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
         Left            =   -73260
         TabIndex        =   17
         Text            =   "0"
         Top             =   480
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BUSCAR"
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
         Left            =   -66780
         TabIndex        =   16
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtEstanteriaDesde 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1980
         TabIndex        =   8
         Text            =   "4"
         Top             =   1980
         Width           =   1035
      End
      Begin VB.TextBox txtVerticalDesde 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1980
         TabIndex        =   7
         Text            =   "1"
         Top             =   2820
         Width           =   1035
      End
      Begin VB.TextBox txtEstado 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1980
         TabIndex        =   6
         Text            =   "1"
         Top             =   3240
         Width           =   1035
      End
      Begin VB.TextBox txtEstanteriaHasta 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   5
         Text            =   "4"
         Top             =   1980
         Width           =   1035
      End
      Begin VB.TextBox txtHorinzontalHasta 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   4
         Text            =   "22"
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox txtVerticalHasta 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   3
         Text            =   "5"
         Top             =   2820
         Width           =   1035
      End
      Begin VB.CommandButton cmdCrear 
         Caption         =   "Crear"
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
         Left            =   3300
         TabIndex        =   2
         Top             =   3420
         Width           =   1035
      End
      Begin VB.ComboBox cboTipoCaja 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Contenedor.frx":0054
         Left            =   1860
         List            =   "Contenedor.frx":005E
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   960
         Width           =   2475
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7155
         Left            =   -74820
         TabIndex        =   18
         Top             =   960
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   12621
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   16
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
      Begin VB.Label Label9 
         Caption         =   "ESTADO: "
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
         Left            =   -68700
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ESTANTERIA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74520
         TabIndex        =   19
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "Estanteria : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   15
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Horizontal :"
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
         Left            =   300
         TabIndex        =   14
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Vertical :"
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
         Left            =   300
         TabIndex        =   13
         Top             =   2820
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Estado :"
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
         Left            =   300
         TabIndex        =   12
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Desde"
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
         Left            =   1980
         TabIndex        =   11
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Hasta"
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
         Left            =   3240
         TabIndex        =   10
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Caja :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1020
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmContenedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrear_Click()
Dim Estanteria As Integer
Dim Horizontal As Integer
Dim Vertical As Integer
Dim NRO_ESTANTE  As Integer
Dim estado As Integer
Dim sSQL As String
Dim Modulo_V As String
Dim Modulo_H As String
Dim Modulo As Long
Dim RS As New ADODB.Recordset
On Error GoTo salir
   Dim ConBasa As New ADODB.Connection
    ConBasa.Open strConBasa
     For Estanteria = txtEstanteriaDesde To txtEstanteriaHasta
        For Vertical = txtVerticalDesde To txtVerticalHasta
            For Horizontal = txtHorinzontalDesde To txtHorinzontalHasta
                Select Case Horizontal
                    Case 16, 17, 18
                        NRO_ESTANTE = 7
                    Case 19, 20, 21
                        NRO_ESTANTE = 8
                End Select
                If cboTipoCaja.Text = "Chica" Then
                  estado = 1
                Else
                  estado = EstadoTipoCaja(Vertical, False)
                End If
                estado = 1
                Select Case Vertical
                    Case 1, 2, 3, 4, 5, 6, 7, 8
                        Modulo_V = 1
                    Case 9, 10, 11, 12, 13, 14, 15, 16
                       Modulo_V = 2
                    Case 17, 18, 19, 20, 21, 22, 23, 24
                       Modulo_V = 3
                    Case 25, 26, 27, 28, 29, 30, 31, 32
                       Modulo_V = 4
                    Case 33, 34, 35, 36, 37, 38, 39, 40
                       Modulo_V = 5
                    Case 41, 42, 43, 44, 45, 46, 47, 48
                       Modulo_V = 6
                    Case 49, 50, 51, 52, 53, 54, 55, 56
                       Modulo_V = 7
                    Case 57, 58, 59, 60, 61, 62, 63, 64
                       Modulo_V = 8
                    Case 65, 66, 67, 68, 69, 70, 71, 72
                       Modulo_V = 9
                    Case 73, 74, 75, 76, 77, 78, 79, 80
                       Modulo_V = 10
                    Case 81, 82, 83, 84, 85, 86, 87, 88
                       Modulo_V = 11
                    Case 89, 90, 91, 92, 93, 94, 95, 96
                       Modulo_V = 12
                    Case 97, 98, 99, 100, 101, 102, 103, 104
                       Modulo_V = 13
                     Case 105, 106, 107, 108, 109, 110, 111, 112
                       Modulo_V = 14
                     Case 113, 114, 115, 116, 117, 118, 119, 120
                       Modulo_V = 15
                     Case 121, 122, 123, 124, 125, 126, 127, 128
                       Modulo_V = 16
                     Case 129, 130, 131, 132, 133, 134, 135, 136
                       Modulo_V = 17
                     Case 137, 138, 139, 140, 141, 142, 143, 144
                       Modulo_V = 18
                     Case 145, 146, 147, 148, 149, 150, 151, 152
                       Modulo_V = 19
                     Case 152, 153, 154, 155, 156, 157, 158, 159
                       Modulo_V = 20
                     Case 160, 161, 162, 163, 164, 165, 166, 167
                        Modulo_V = 21
                     Case 168, 169, 170, 171, 172, 173, 174, 175
                        Modulo_V = 22
                     Case 176, 177, 178, 179, 180, 181, 182, 183
                        Modulo_V = 23
                     Case 184, 185, 186, 187, 188, 189, 190, 191
                        Modulo_V = 24
                    Case 191, 192, 193, 194, 195, 196, 197, 198
                        Modulo_V = 25
                    Case 199, 200, 201, 202, 203, 204, 205, 206
                        Modulo_V = 26
                    Case 207, 208, 209, 210, 211, 212, 213, 214
                        Modulo_V = 27
                End Select
                Select Case Horizontal
                    Case 1, 2, 3, 4, 5
                        Modulo_H = 1
                    Case 6, 7, 8, 9, 10
                        Modulo_H = 2
                    Case 11, 12, 13, 14, 15
                        Modulo_H = 3
                    Case 16, 17, 18, 19, 20
                        Modulo_H = 4
                    Case 21, 22, 23, 24, 25
                        Modulo_H = 5
                End Select
                Rem luisestanteria
                sSQL = " SELECT ESTANTERIA, HORIZONTAL, VERTICAL "
                sSQL = sSQL & vbCrLf & " From CONTENEDOR "
                sSQL = sSQL & vbCrLf & " Where ESTANTERIA = " & Estanteria
                sSQL = sSQL & vbCrLf & " And Horizontal = " & Horizontal
                sSQL = sSQL & vbCrLf & " And Vertical = " & Vertical
                
                Set RS = New ADODB.Recordset
                RS.Open sSQL, strConBasa
                
                If RS.EOF Then
                    Modulo = Str(Estanteria) & Modulo_V & Modulo_H & 1
                    sSQL = "INSERT INTO CONTENEDOR "
                    sSQL = sSQL & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
                    sSQL = sSQL & vbCrLf & " NRO_ESTANTE, ESTADO, MODULO_V, MODULO_H , FECHA_CREACION )"
                    sSQL = sSQL & vbCrLf & " VALUES (" & Estanteria & "," & Horizontal & "," & Vertical & ",1," & NRO_ESTANTE & "," & estado & "," & Modulo_V & "," & Modulo_H & ", '16/05/2011' )"
                    sSQL = sSQL & vbCrLf & ""
                    ExecutarSql (sSQL)
                End If
                 
            Next
        Next
    Next
    MsgBox " terminado"
    Exit Sub
    
salir:
    
    MsgBox "error"
    
End Sub
Public Function EstadoTipoCaja(Vertical As Integer, TipoCajaChica As Boolean) As Integer
EstadoTipoCaja = 1
If TipoCajaChica = True Then
Else
Select Case Vertical
Case 7, 8, 15, 16, 23, 24, 32, 31, 39, 40, 48, 47, 55, 56, 64, 63
    EstadoTipoCaja = 0
End Select
End If

End Function

Private Sub Command1_Click()
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
Dim Sql As String


Sql = " SELECT     ID_CONTENEDOR, ESTANTERIA, VERTICAL, HORIZONTAL, ADELANTE_ATRAS, ESTADO, COD_CLIENTE, NRO_CAJA"
Sql = Sql & " From CONTENEDOR "
Sql = Sql & " Where "
If txtEstanteriaDesde1.Text = 0 Then
    MsgBox "Ingrese Estanteria"
    Exit Sub
Else
    If txtEstanteriaHasta2.Text = 0 Then
       Sql = Sql & " Estanteria = " & txtEstanteriaDesde1.Text
    Else
       Sql = Sql & " Estanteria Between " & txtEstanteriaDesde1.Text & " and " & txtEstanteriaHasta2.Text
    End If
End If

If txtEstado1.Text <> "" Then
   Sql = Sql & " AND ESTADO = " & txtEstado1.Text
End If

Sql = Sql & " ORDER BY ESTANTERIA,  VERTICAL, HORIZONTAL"

RS.Open Sql, ConActiva, 2, 3

Set DataGrid1.DataSource = RS.DataSource
Dim I As Integer
For I = 0 To DataGrid1.Columns.Count - 1
    DataGrid1.Columns.Item(I).Locked = True
Next

DataGrid1.Columns.Item("ESTADO").Locked = False


End Sub

Private Sub Command2_Click()
CopiarDatosGrilla DataGrid1
End Sub

Private Sub Command3_Click()
        Dim Estanteria As Integer
        Dim Horizontal As Integer
        Dim Vertical As Integer
        Dim NRO_ESTANTE  As Integer
        Dim estado As Integer
        Dim sSQL As String
        Dim Modulo_V As String
        Dim Modulo_H As String
        Dim Modulo As Long
        Dim RS As ADODB.Recordset
        Dim Sql As String
        
           Dim ConBasa As New ADODB.Connection
                ConBasa.Open strConBasa
             
             For Estanteria = txtEstanteriaDesde To txtEstanteriaHasta
                For Vertical = txtVerticalDesde To txtVerticalHasta
                    For Horizontal = txtHorinzontalDesde To txtHorinzontalHasta
                        If cboTipoCaja.Text = "Chica" Then
                          estado = 1
                        Else
                          estado = EstadoTipoCaja(Vertical, False)
                        End If
                        
                         estado = 1
                        
                        Select Case Vertical
                        Case 1, 2, 3, 4, 5, 6, 7, 8
                            Modulo_V = 1
                        Case 9, 10, 11, 12, 13, 14, 15, 16
                           Modulo_V = 2
                        Case 17, 18, 19, 20, 21, 22, 23, 24
                           Modulo_V = 3
                        Case 25, 26, 27, 28, 29, 30, 31, 32
                           Modulo_V = 4
                        Case 33, 34, 35, 36, 37, 38, 39, 40
                           Modulo_V = 5
                        Case 41, 42, 43, 44, 45, 46, 47, 48
                           Modulo_V = 6
                        Case 49, 50, 51, 52, 53, 54, 55, 56
                           Modulo_V = 7
                        Case 57, 58, 59, 60, 61, 62, 63, 64
                           Modulo_V = 8
                        Case 65, 66, 67, 68, 69, 70, 71, 72
                           Modulo_V = 9
                        Case 73, 74, 75, 76, 77, 78, 79, 80
                           Modulo_V = 10
                        Case 81, 82, 83, 84, 85, 86, 87, 88
                           Modulo_V = 11
                        Case 89, 90, 91, 92, 93, 94, 95, 96
                           Modulo_V = 12
                        Case 97, 98, 99, 100, 101, 102, 103, 104
                           Modulo_V = 13
                         Case 105, 106, 107, 108, 109, 110, 111, 112
                           Modulo_V = 14
                         Case 113, 114, 115, 116, 117, 118, 119, 120
                           Modulo_V = 15
                         Case 121, 122, 123, 124, 125, 126, 127, 128
                           Modulo_V = 16
                         Case 129, 130, 131, 132, 133, 134, 135, 136
                           Modulo_V = 17
                         Case 137, 138, 139, 140, 141, 142, 143, 144
                           Modulo_V = 18
                         Case 145, 146, 147, 148, 149, 150, 151, 152
                            Modulo_V = 19
                         Case 153, 154, 155, 156, 157, 158, 159, 160
                            Modulo_V = 20
                         Case 161, 162, 163, 164, 165, 166, 167, 168
                            Modulo_V = 21
                         Case 169, 170, 171, 172, 173, 174, 175, 176
                            Modulo_V = 22
                          Case 177, 178, 179, 180, 181, 182, 183, 184
                            Modulo_V = 23
                          Case 185, 186, 187, 188, 189, 190, 191, 192
                            Modulo_V = 24
                          Case 193, 194, 195, 196, 197, 198, 199, 200
                            Modulo_V = 25
                          Case 201, 202, 203, 204, 205, 206, 207, 208
                            Modulo_V = 26
                        End Select
                       Select Case Horizontal
                       Case 1, 2, 3, 4, 5
                            Modulo_H = 1
                       Case 6, 7, 8, 9, 10
                            Modulo_H = 2
                       Case 11, 12, 13, 14, 15
                            Modulo_H = 3
                       Case 16, 17, 18, 19, 20
                            Modulo_H = 4
                       Case 21, 22, 23, 24, 25
                            Modulo_H = 5
                       End Select
                       Set RS = New ADODB.Recordset
                       
                       Sql = " Select * FROM CONTENEDOR where "
                       Sql = Sql & " Estanteria  = " & Estanteria
                       Sql = Sql & " AND Horizontal  = " & Horizontal
                       Sql = Sql & " AND Vertical  = " & Vertical
                       
                        RS.Open Sql, strConBasa
                       
                       If RS.EOF Then
                            Modulo = Str(Estanteria) & Modulo_V & Modulo_H & 1
                            sSQL = "INSERT INTO CONTENEDOR "
                            sSQL = sSQL & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
                            sSQL = sSQL & vbCrLf & " NRO_ESTANTE, ESTADO,MODULO_V, MODULO_H)"
                            sSQL = sSQL & vbCrLf & " VALUES (" & Estanteria & "," & Horizontal & "," & Vertical & ", 1 ," & NRO_ESTANTE & "," & estado & "," & Modulo_V & "," & Modulo_H & ")"
                            ExecutarSql (sSQL)
                       End If
                    Next
                Next
            Next
End Sub

Private Sub Form_Load()
cboTipoCaja.ListIndex = 0
End Sub

