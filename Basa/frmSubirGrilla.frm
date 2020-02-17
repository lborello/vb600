VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSubirGrilla 
   Caption         =   "Subir Grilla"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdFinal 
      Height          =   1155
      Left            =   0
      TabIndex        =   2
      Top             =   2340
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   2037
      _Version        =   393216
      Cols            =   10
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   6180
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdPasar 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   10
   End
End
Attribute VB_Name = "frmSubirGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim rs As ADODB.Recordset
Dim Sql As String


Sql = " SELECT     ID, PASO, LEIDO"
Sql = Sql & " From basasql.dbo.Paso_txt"
Sql = Sql & " ORDER BY ID"
Dim a As VbFileAttribute
Dim Paso As String

Set rs = New ADODB.Recordset
rs.Open Sql, strConBasa
Do While Not rs.EOF
Paso = Replace(rs!Paso, "¢", "ó")
Paso = Replace(Paso, "¥", "ñ")

    LeerArchivo Paso, FileSystem.FileDateTime(Paso)
 



    rs.MoveNext
    Loop
    


End Sub

Public Sub LeerArchivo(Paso As String, fecha As String)
On Error GoTo salir:
    Open Paso For Input As #1
    Dim P As Integer
    Dim con As New ADODB.Connection
    
    Do Until EOF(1)
        Line Input #1, VarTexto
        If VarTexto <> "" Then
            grdPasar.AddItem VarTexto
            
            
            If IsNumeric(grdPasar.TextMatrix(2, 2)) Then
            
            grdFinal.AddItem VarTexto & vbTab & fecha & vbTab & Paso
            
                grdPasar.Rows = 2
            grdPasar.Clear
            
            Else
            grdPasar.Rows = 2
            grdPasar.Clear
            End If
            
            
        End If
    Loop
    
        Dim Sql As String
        Dim Orden As String
        Dim Caja As Double
        Dim DATO1 As String
        Dim DATO2 As String
        Dim DATO3 As String
        Dim DATO4 As String
        Dim DATO5 As String
        Dim DATO6 As String
        Dim DATO7 As String
        Dim FK_ARCHIVO As Long

    If grdFinal.Rows > 2 Then
    
    
        For i = 2 To grdFinal.Rows - 1
            Orden = ""
            Caja = 0
            DATO1 = ""
            DATO2 = ""
            DATO3 = ""
            DATO4 = ""
            DATO5 = ""
            DATO6 = ""
            DATO7 = ""
            FK_ARCHIVO = ""
            If Trim(grdFinal.TextMatrix(i, 1)) <> "" Then
                Orden = "'" & Trim(grdFinal.TextMatrix(i, 1)) & "'"
            Else
                Orden = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 2)) <> "" Then
                If IsNumeric(grdFinal.TextMatrix(i, 2)) Then
                     If Len(grdFinal.TextMatrix(i, 2)) > 12 Then
                
                    Caja = 0
                    Else
                    
                    Caja = grdFinal.TextMatrix(i, 2)
                    End If
                    
                Else
                Caja = 0
                End If
            Else
                Caja = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 3)) <> "" Then
                DATO1 = "'" & Trim(grdFinal.TextMatrix(i, 3)) & "'"
            Else
                DATO1 = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 4)) <> "" Then
                DATO2 = "'" & Trim(grdFinal.TextMatrix(i, 4)) & "'"
            Else
                DATO2 = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 5)) <> "" Then
                DATO3 = "'" & Trim(grdFinal.TextMatrix(i, 5)) & "'"
            Else
                DATO3 = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 6)) <> "" Then
                DATO4 = "'" & Trim(grdFinal.TextMatrix(i, 6)) & "'"
            Else
                DATO4 = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 7)) <> "" Then
                DATO5 = "'" & Trim(grdFinal.TextMatrix(i, 7)) & "'"
            Else
                DATO5 = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 8)) <> "" Then
                DATO6 = "'" & Trim(grdFinal.TextMatrix(i, 8)) & "'"
            Else
                DATO6 = "NULL"
            End If
            If Trim(grdFinal.TextMatrix(i, 9)) <> "" Then
                DATO7 = "'" & Trim(grdFinal.TextMatrix(i, 9)) & "'"
            Else
                DATO7 = "NULL"
            End If
            DATO8 = "NULL"
            
            If Trim(grdFinal.TextMatrix(i, 10)) <> "" Then
                FK_ARCHIVO = Trim(grdFinal.TextMatrix(i, 10))
            Else
                FK_ARCHIVO = 0
            End If
            
            
            
            
            Sql = " Insert "
            Sql = Sql & " Into basasql.dbo.LECTURA_CONTROL_ALSINA("
            Sql = Sql & vbCrLf & " Orden,"
            Sql = Sql & vbCrLf & " Caja,"
            Sql = Sql & vbCrLf & " DATO1,"
            Sql = Sql & vbCrLf & " DATO2, "
            Sql = Sql & vbCrLf & " DATO3, "
            Sql = Sql & vbCrLf & " DATO4,"
            Sql = Sql & vbCrLf & " DATO5,"
            Sql = Sql & vbCrLf & " DATO6,"
            Sql = Sql & vbCrLf & " DATO7,"
            Sql = Sql & vbCrLf & " FK_ARCHIVO)"
            Sql = Sql & vbCrLf & " VALUES("
            Sql = Sql & vbCrLf & Orden
            Sql = Sql & vbCrLf & "," & Caja
            Sql = Sql & vbCrLf & "," & DATO1
            Sql = Sql & vbCrLf & "," & DATO2
            Sql = Sql & vbCrLf & "," & DATO3
            Sql = Sql & vbCrLf & "," & DATO4
            Sql = Sql & vbCrLf & "," & DATO5
            Sql = Sql & vbCrLf & "," & DATO6
            Sql = Sql & vbCrLf & "," & DATO7
            Sql = Sql & vbCrLf & "," & FK_ARCHIVO & ")"
            Set con = New ADODB.Connection
    con.ConnectionString = strConBasa
   con.Open
            con.Execute Sql
        Next
    End If
    
    
    grdFinal.Rows = 2
    grdFinal.Clear
    grdFinal.Refresh
 Close #1
  Exit Sub
salir:
MsgBox Err.Description
 Close #1
End Sub
