VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmAndesmar 
   Caption         =   "Andesmar"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   12990
   Begin VB.CommandButton Command11 
      Caption         =   "ExportarOrdenes"
      Height          =   315
      Left            =   8880
      TabIndex        =   22
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ExportarNUevo"
      Height          =   315
      Left            =   8700
      TabIndex        =   21
      Top             =   180
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   10320
      Width           =   1155
   End
   Begin VB.CommandButton cmdOrdenesAndesmar 
      Caption         =   "Orden Andesmar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   1320
      Width           =   1875
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   10320
      Width           =   1395
   End
   Begin VB.CommandButton cmdDansu 
      Caption         =   "Damsu Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   720
      Width           =   1875
   End
   Begin VB.CommandButton cmdDansuDigito 
      Caption         =   "Sin Digito Verificador Damsu"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   720
      Width           =   2475
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   720
      Width           =   1875
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Barra Error"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   1875
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCorregirArchivos 
      Caption         =   "Corregir Archivos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   720
      Width           =   1875
   End
   Begin VB.TextBox txtCaja 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10800
      TabIndex        =   9
      Top             =   1380
      Width           =   1455
   End
   Begin VB.TextBox txtPasoGuia 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Text            =   "I:\0403 -ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS"
      Top             =   1320
      Width           =   6195
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cmdExportTxt"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   10320
      Width           =   1395
   End
   Begin VB.CommandButton cmdControlLotes 
      Caption         =   "Control Lotes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   1875
   End
   Begin VB.CommandButton cmdBarra 
      Caption         =   "Barra Todos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1875
   End
   Begin MSComDlg.CommonDialog otpAbrirArchivo 
      Left            =   180
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCambioNombre 
      Caption         =   "Cambio de Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   120
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Indexsar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4380
      TabIndex        =   2
      Top             =   120
      Width           =   1875
   End
   Begin MSDataGridLib.DataGrid grdIndexarImagenes 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   1860
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
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
   Begin Controles.ctlVerImagenes ctlVerImagenes1 
      Height          =   5415
      Left            =   165
      TabIndex        =   0
      Top             =   4620
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9551
   End
   Begin VB.Label Label1 
      Caption         =   "Caja:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9840
      TabIndex        =   10
      Top             =   1440
      Width           =   675
   End
End
Attribute VB_Name = "frmAndesmar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGrilla As ADODB.Recordset
Dim Pasar As Boolean

Private Sub cmdBarra_Click()
    Dim Sql As String
        Set rsGrilla = New ADODB.Recordset
        Pasar = False
        Sql = "SELECT ID, Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, BARRA,  BatchPgDta"
        Sql = Sql & " From TELEFORM_BARRA"
        Sql = Sql & " WHERE ( BatchNo = " & InputBox("Ingrese el lote") & ")"
       Rem Sql = Sql & " WHERE   (CONVERT(char, BARRA) LIKE '4372000126%') and BatchNo = 2412"
        Sql = Sql & "  ORDER BY BARRA DESC"
        rsGrilla.CursorLocation = adUseClient
        rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic
        Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
        grdIndexarImagenes.ReBind
        grdIndexarImagenes.Refresh
End Sub

Private Sub cmdCambioNombre_Click()
Dim direcOrig(400) As String
Dim direcFin(400) As String

Dim i As Integer
Dim Caja As Long
Dim sFolderPath As String
Dim sArchivo As String
Dim PasoFinal As String
Dim PasoTeleform As String
Dim PasoDig As String
Dim PasoDirSubDir As String


If Trim(txtCaja.Text) <> "" Then
    Caja = txtCaja.Text
Else
    MsgBox "Ingrese la caja"
    Exit Sub
End If



   PasoDig = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\DIGITALIZADAS\"
    sArchivo = Dir(PasoDig & Caja & "\*", vbDirectory)
    
    If sArchivo = "" Then
        MsgBox "No existe la caja" & Caja & " en " & PasoDig
        Exit Sub
    End If
    
    Do While sArchivo <> ""
        
        If Len(sArchivo) > 8 Then
        direcOrig(i) = sArchivo
        direcFin(i) = Mid(sArchivo, 1, 7) & Format(Mid(sArchivo, 8), "000")
        Else
        direcOrig(i) = ""
        direcFin(i) = ""
        End If
        i = i + 1
        sArchivo = Dir
    Loop

PasoTeleform = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM"

If Dir(PasoTeleform & "\" & Caja, vbDirectory) = "" Then
 FileSystem.MkDir PasoTeleform & "\" & Caja
 PasoFinal = PasoTeleform & "\" & Caja
Else
    PasoFinal = PasoTeleform & "\" & Caja
End If

    Dim ArchivoNombreFinal As String
        For i = 0 To 300
           If Trim(direcOrig(i)) <> "" Then
              If Len(Trim(direcOrig(i))) > 2 Then
                PasoDirSubDir = PasoDig & Caja & "\" & direcOrig(i)
                sArchivo = Dir(PasoDirSubDir & "\*.tif")
                Do While sArchivo <> ""
                   ArchivoNombreFinal = Mid(sArchivo, 1, Len(sArchivo) - 4)
                   ArchivoNombreFinal = Format(ArchivoNombreFinal, "000")
                   FileCopy PasoDirSubDir & "\" & sArchivo, PasoFinal & "\" & direcFin(i) & ArchivoNombreFinal & ".TIF"
                   sArchivo = Dir
                Loop
              End If
            End If
        Next
    MsgBox "Terminados"


End Sub

Private Sub cmdConDigitoVerificador_Click()


Dim Sql As String
 Dim rs As New ADODB.Recordset
  Dim rsLega As New ADODB.Recordset

Sql = " SELECT     ID, BatchPgDta, BARRA, LEN(BARRA) AS Expr1, BatchNo"
Sql = Sql & "  From TELEFORM_BARRA"
Sql = Sql & "  WHERE     (BatchNo IN (1348, 1350)) AND (LEN(BARRA) = 8)"
Sql = Sql & "  ORDER BY BARRA"


rs.Open Sql, strConBasa

Do While Not rs.EOF
        Sql = " SELECT TOP (10) ID_LEGAJO, ETIQUETA, DIGITO_VERIFICADOR"
        Sql = Sql & " From basasql.dbo.LEGAJOS"
        Sql = Sql & " Where ID_LEGAJO = " & Mid(rs!BARRA, 1, 7)
        Sql = Sql & " And Digito_Verificador = " & Mid(rs!BARRA, 8, 1)
       
        
        

    Set rsLega = New ADODB.Recordset
    
        rsLega.Open Sql, strConBasa
    
        If Not rsLega.EOF Then
            Sql = "  Update TELEFORM_BARRA"
            Sql = Sql & " Set BARRA = " & rsLega!Etiqueta
            Sql = Sql & " Where ID = " & rs!ID
            ExecutarSql Sql
        End If
        
        rs.MoveNext
Loop

MsgBox "tERMINADO"

End Sub

Private Sub cmdControlLotes_Click()
    Dim Sql As String
        Set rsGrilla = New ADODB.Recordset
        Pasar = False
        Sql = "SELECT     SUBSTRING(BatchPgDta, 1, 10) AS Lote, COUNT(*) AS Cantidad"
        Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
        Sql = Sql & " GROUP BY SUBSTRING(BatchPgDta, 1, 10), BatchNo"
        Sql = Sql & " Having (BatchNo = " & InputBox("Ingrese el lote") & ")"
        Sql = Sql & " ORDER BY Lote"
        rsGrilla.CursorLocation = adUseClient
        rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic
        Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
        grdIndexarImagenes.ReBind
        grdIndexarImagenes.Refresh

End Sub

Private Sub cmdCorregirArchivos_Click()
Dim PasoImagenesConError As String
Dim PasoImagenesConBien As String
Dim PasoImagenesEx As String
Dim sArchivo  As String
Dim ArchivoNombreFinal  As String
Dim Caja As String

PasoImagenesConError = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS\Archivos Con Error"
PasoImagenesConBien = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS\Archivos Con Solucion"
PasoImagenesEx = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS"

Dir (PasoImagenesConError)

  sArchivo = Dir(PasoImagenesConError & "\*.PDF")
    Do While sArchivo <> ""
        ArchivoNombreFinal = sArchivo
        FileSystem.FileCopy PasoImagenesEx & "\LOTE" & Mid(sArchivo, 1, 7) & "\IMAGENES\" & sArchivo, PasoImagenesConBien & "\" & sArchivo
        sArchivo = Dir
    Loop
    
    

End Sub

Private Sub cmdDansu_Click()

    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Caja As String



Sql = " SELECT  ID , BARRA , NombreArchivoNumero "
Sql = Sql & " From TELEFORM_BARRA"
Sql = Sql & " WHERE  BatchNo  IN (" & InputBox("Ingrese Lote") & ")"
Sql = Sql & "  AND (BARRA <> 0)"
Sql = Sql & " ORDER BY ID "


'        Sql = "  SELECT     TELEFORM_BARRA.ID, TELEFORM_BARRA.BARRA, TELEFORM_BARRA.NombreArchivoNumero, LEGAJOS.DIGITO_VERIFICADOR"
'        Sql = Sql & " FROM         TELEFORM_BARRA INNER JOIN"
'        Sql = Sql & " LEGAJOS ON TELEFORM_BARRA.BARRA = LEGAJOS.ETIQUETA"
'        Sql = Sql & " WHERE     (TELEFORM_BARRA.BatchNo IN (1410,1411,1412)) AND (TELEFORM_BARRA.BARRA <> 0)"
'        Sql = Sql & " ORDER BY TELEFORM_BARRA.ID"
        
       
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
            Caja = Mid(rs!NombreArchivoNumero, 1, 7)
            If Dir("D:\Dansu\" & Caja & "\" & rs!NombreArchivoNumero & ".tif") <> "" Then
                FileCopy "D:\Dansu\" & Caja & "\" & rs!NombreArchivoNumero & ".tif", "D:\Dansu\NombreCodigo\" & Caja & "\" & rs!BARRA & ".tif"
            Else
                Debug.Print rs!NombreArchivoNumero & ".tif"
            End If
            
            rs.MoveNext
        Loop




End Sub

Private Sub cmdDansuDigito_Click()
 Dim Sql As String
 Dim rs As New ADODB.Recordset
'
'
'sql = " SELECT     ID, BatchPgDta, BARRA"
'sql = sql & " From TELEFORM_BARRA"
'sql = sql & " WHERE     (BatchNo IN (1348, 1350)) AND (LEN(BARRA) = 13)"
'sql = sql & " ORDER BY ID"
'
'
'rs.Open sql, strConBasa
'
'Do While Not rs.EOF
'    sql = "  Update TELEFORM_BARRA "
'    sql = sql & " Set BARRA = " & Mid(rs!BARRA, 6, 7)
'    sql = sql & " Where ID = " & rs!ID
'    ExecutarSql sql
'    rs.MoveNext
'Loop



        Sql = " SELECT TELEFORM_BARRA.ID, TELEFORM_BARRA.BatchPgDta, "
        Sql = Sql & " TELEFORM_BARRA.BARRA, LEN(TELEFORM_BARRA.BARRA) AS Expr1, LEGAJOS.ETIQUETA,"
        Sql = Sql & " TELEFORM_BARRA.BatchNo , LEGAJOS.DIGITO_VERIFICADOR "
        Sql = Sql & " FROM TELEFORM_BARRA INNER JOIN "
        Sql = Sql & " LEGAJOS ON TELEFORM_BARRA.BARRA = LEGAJOS.ID_LEGAJO "
        Sql = Sql & " WHERE TELEFORM_BARRA.BatchNo IN (" & InputBox("Ingrese Lote") & ")"
        Sql = Sql & " AND (LEN(TELEFORM_BARRA.BARRA) = 7)"
        Sql = Sql & " ORDER BY TELEFORM_BARRA.BARRA"
        rs.Open Sql, strConBasa
    
    Do While Not rs.EOF
        Sql = "  Update TELEFORM_BARRA"
        Sql = Sql & " Set BARRA = " & rs!Etiqueta & rs!Digito_Verificador
        Sql = Sql & " Where ID = " & rs!ID
        ExecutarSql Sql
        rs.MoveNext
    Loop
    
    
    
    





End Sub

Private Sub cmdOrdenesAndesmar_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Nombre_Archivo As String
'    Sql = " SELECT     BatchNo, Suspense_File, BatchPgDta, SUBSTRING(BatchPgDta, 1, 7) AS Expr1, barra"
'    Sql = Sql & " From  TELEFORM_BARRA "
'    Sql = Sql & " WHERE BatchNo IN ( )"
'
'
'Sql = "SELECT     BatchNo, Suspense_File, BatchPgDta, SUBSTRING(BatchPgDta, 1, 7) AS Expr1, barra"
' Sql = Sql & " From TELEFORM_BARRA"
' Sql = Sql & " WHERE     (BatchNo IN (" & InputBox("Ingrese el numero Batch") & ") AND (BARRA > 200)"
' Sql = Sql & " ORDER BY BARRA"

Sql = "SELECT   BatchNo, Suspense_File, BatchPgDta, SUBSTRING(BatchPgDta, 1, 7) AS Expr1, NRO_DESDE2 as BARRA "
Sql = Sql & " From TELEFORM_BARRA"
Sql = Sql & " Where (BatchNo = 2167)  "
Sql = Sql & " ORDER BY BARRA"



Sql = "SELECT id,  BatchNo, Suspense_File, BatchPgDta, SUBSTRING(BatchPgDta, 1, 7) AS Expr1,  BARRA "
Sql = Sql & " From TELEFORM_BARRA"
Sql = Sql & " Where (BatchNo in(2616))  "
Sql = Sql & " ORDER BY BARRA"



rs.Open Sql, strConBasa

Do While Not rs.EOF
    Nombre_Archivo = Format(rs!BARRA, "00000000") & "_" & rs!ID
   FileCopy "\\PCTELEMEMO1" & Mid(rs!Suspense_File, 3), "D:\AndesmarTerminados\" & Nombre_Archivo & ".tif"
   rs.MoveNext

Loop




End Sub

Private Sub Command1_Click()
    Dim Sql As String
        Set rsGrilla = New ADODB.Recordset
        Pasar = False
        Dim NumeroLote As Long
        Dim rs As New ADODB.Recordset
        
        NumeroLote = InputBox("Ingrese el lote")
        
        
        Sql = "SELECT     NRO_DESDE2, ID, NRO_DESDE"
        Sql = Sql & " From TELEFORM_DIGITAL"
        Sql = Sql & " Where (BatchNo = " & NumeroLote & ") And Len(NRO_DESDE2) < 5"
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
            Sql = " Update TELEFORM_DIGITAL"
            If IsNumeric(rs!NRO_DESDE) And Len(rs!NRO_DESDE) > 5 Then
                Sql = Sql & " Set NRO_DESDE2 =" & Trim(rs!NRO_DESDE)
            Else
            If IsNull(rs!NRO_DESDE2) Then
                Sql = Sql & " Set NRO_DESDE2 =0"
           Else
           Sql = Sql & " Set NRO_DESDE2 =" & rs!NRO_DESDE2
            End If
            End If
            
            Sql = Sql & " Where ID =" & rs!ID
            ExecutarSql Sql
            
            rs.MoveNext
        Loop
        
        
        Sql = " SELECT      ID, Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, NRO_DESDE2,  BatchPgDta "
        Sql = Sql & " From TELEFORM_DIGITAL"
        Sql = Sql & " WHERE    (BatchNo = " & NumeroLote & ")"
        
        Sql = Sql & "  ORDER BY NRO_DESDE2 DESC"
        
        

        
        
        
        
        
        rsGrilla.CursorLocation = adUseClient
        rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic
        Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
        grdIndexarImagenes.ReBind
        grdIndexarImagenes.Refresh
End Sub

Private Sub Command10_Click()
Dim Sql As String
    Dim Caja As Long
    Dim Encabezado As String
    Dim NombreArchivoImagenPDF As String
    Dim NombreArchivoImagenTIF As String
    Dim NombreArchivoTxt As String
    Dim LoteNombreArchivoTxt As String
    Dim PasoInicio As String
    Dim PasoImagenes As String
    Dim PasoExpImagenes  As String
    Dim PasoExp As String
    Dim Datos As String
    Dim CantImagenes As Long
    Dim ControlXML As String
    Dim PasoExpRaiz As String
    Dim CANTIDAD_IMAGEN As Integer
    Dim ArchivoNoEncontrado As String
    
    Dim rs As New ADODB.Recordset
    Dim Strlotes As String
    
    Set rsGrilla = New ADODB.Recordset
    Pasar = False
    
    Caja = InputBox(" INGRESE LA CAJA ")
    Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " FROM  DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " Where DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = " & Caja
    Sql = Sql & " AND DOCUMENTOS_DIGITALES.ESTADO = 'LISTA PARA EXPORTAR'"
    Sql = Sql & " GROUP BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
    Do While Not rs.EOF
        Strlotes = Strlotes & "," & rs!ID_DOCUMENTOS_DIGITALES_LOTE
        rs.MoveNext
    Loop
    
    
    
    
   Sql = "  SELECT        DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN,"
   Sql = Sql & "                       DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.ESTADO, DOCUMENTOS_DIGITALES.CONVERSION_TIPO_ARCHIVO,"
    Sql = Sql & "                      DOCUMENTOS_DIGITALES.DIRECTORIO_PASO ,  DOCUMENTOS_DIGITALES.CANTIDAD_IMAGENES ,  DOCUMENTOS_DIGITALES.ID  AS IDIMAGEN"
Sql = Sql & "  FROM            DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & "                          DOCUMENTOS_DIGITALES ON"
Sql = Sql & "                          DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where  DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = " & Caja & " "
 Sql = Sql & "  AND ESTADO = 'LISTA PARA EXPORTAR'"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
    
    Rem RS.CursorLocation = adUseClient
    
   Set rs = New ADODB.Recordset
   
       rs.Open Sql, strConBasa
  
    If rs.EOF Then
        MsgBox "nO EXISTE LA CAJA"
        Exit Sub
        
    End If
    
        
        
        LoteNombreArchivoTxt = "LOTE" & Caja
        PasoImagenes = "\\222.15.19.251\ImagenesPDF\"
       Rem  PasoExp = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS" & "\" & "LOTE" & Caja
       PasoExp = "D:\EXPORTADAS" & "\" & "LOTE" & Caja
    If Dir(PasoExp, vbDirectory) = "" Then
        FileSystem.MkDir PasoExp
        PasoExpRaiz = PasoExp
        PasoExpImagenes = PasoExpRaiz & "\IMAGENES"
        FileSystem.MkDir PasoExpImagenes
     Else
        MsgBox "Los directorios ya existen"
        Exit Sub
    End If
    
    
    
  Rem   I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM
    NombreArchivoTxt = PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".TXT"
    Open NombreArchivoTxt For Append As #1
    Encabezado = Chr(34) & "DocumentFileName" & Chr(34) & "," & Chr(34) & "PageCount" & Chr(34) & "," & Chr(34) & "Guia" & Chr(34) & "," & Chr(34) & "Fecha" & Chr(34) & "," & Chr(34) & "Sucursal" & Chr(34) & "," & Chr(34) & "Caja" & Chr(34) & "," & Chr(34) & "Lote" & Chr(34)
    Print #1, Encabezado
        Do While Not rs.EOF
            CantImagenes = CantImagenes + 1
            
            Sql = " Update DOCUMENTOS_DIGITALES"
            Sql = Sql & " SET ESTADO ='EXPORTADA'"
            Sql = Sql & " Where ID = " & rs!IDIMAGEN
            ExecutarSql Sql
            
            NombreArchivoImagenPDF = Format(rs!FK_CAJAS, "00000000") & "_" & Format(rs!ID_DOCUMENTOS_DIGITALES_LOTE, "000000") & "_" & Format(rs!IMAGEN_ORIGEN, "0000") & "_" & rs!IDIMAGEN & ".pdf"
         Rem  NombreArchivoImagenPDF = RS!IDIMAGEN & ".pdf"
'            NombreArchivoImagenPDF = Mid(rsGrilla!BatchPgDta, 1, 16) & ".pdf"
            NombreArchivoImagenTIF = Format(rs!FK_CAJAS, "00000000") & "_" & Format(rs!ID_DOCUMENTOS_DIGITALES_LOTE, "000000") & "_" & Format(rs!IMAGEN_ORIGEN, "0000") & "_" & rs!IDIMAGEN & ".pdf"
            If IsNull(rs!Cantidad_Imagenes) Then
                CANTIDAD_IMAGEN = 1
            Else
                CANTIDAD_IMAGEN = rs!Cantidad_Imagenes
            End If
            Datos = Chr(34) & "//" & LoteNombreArchivoTxt & "/IMAGENES/" & NombreArchivoImagenPDF & Chr(34) & "," & Chr(34) & CANTIDAD_IMAGEN & Chr(34) & "," & Chr(34) & rs!NRO_DESDE & Chr(34) & "," & Chr(34) & Format(Now, "DD/MM/YYYY") & Chr(34) & "," & Chr(34) & "MENDOZA" & Chr(34) & "," & Chr(34) & Caja & Chr(34) & "," & Chr(34) & Format(rs!ID_DOCUMENTOS_DIGITALES_LOTE, "000000") & Chr(34)
            Print #1, Datos
             
             
            Rem  ESTA MAL FileCopy "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3), PasoExpImagenes & "\" & NombreArchivoImagenTIF
             
             If Dir(PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".PDF", vbArchive) <> "" Then
                 FileCopy PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".PDF", PasoExpImagenes & "\" & NombreArchivoImagenPDF
                Else
                
                    If Dir("\\222.15.19.251\Imagenes" & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".TIF", vbArchive) <> "" Then
                     FileCopy "\\222.15.19.251\Imagenes" & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".TIF", PasoExpImagenes & "\" & NombreArchivoImagenTIF
                    Else
                      MsgBox "no se encontro " & "\\222.15.19.251\Imagenes" & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".TIF"
                    End If
                
            End If
            
            rs.MoveNext
            Debug.Print CantImagenes
        Loop
    Close #1
    ControlXML = "<Batch>" & vbCr
    ControlXML = ControlXML & " <Statistics>" & vbCr
    ControlXML = ControlXML & "     <DocumentCount>" & CantImagenes & "</DocumentCount>" & vbCr
    ControlXML = ControlXML & " </Statistics>" & vbCr
    ControlXML = ControlXML & "</Batch>" & vbCr
    Open PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".XML" For Append As #1
        Print #1, ControlXML
    Close #1
    
    If ArchivoNoEncontrado <> "" Then
                MsgBox (" NO se enconcotro el archivo")
                
                Clipboard.Clear
               Clipboard.SetText ArchivoNoEncontrado
                
    End If
    
    Sql = " Update DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " SET LOTE_ESTADO ='EXPORTADO'"
    Sql = Sql & " , FECHA_EXPORTACION=" & SysDateMinutoSegundo
    Sql = Sql & "  Where ID_DOCUMENTOS_DIGITALES_LOTE IN( " & Mid(Strlotes, 2) & ")"
    
    ExecutarSql Sql

    
    MsgBox "terminados"
End Sub

Private Sub Command11_Click()
Dim Sql As String
    Dim Caja As Long
    Dim Encabezado As String
    Dim NombreArchivoImagenPDF As String
    Dim NombreArchivoImagenTIF As String
    Dim NombreArchivoTxt As String
    Dim LoteNombreArchivoTxt As String
    Dim PasoInicio As String
    Dim PasoImagenes As String
    Dim PasoExpImagenes  As String
    Dim PasoExp As String
    Dim Datos As String
    Dim CantImagenes As Long
    Dim ControlXML As String
    Dim PasoExpRaiz As String
    Dim CANTIDAD_IMAGEN As Integer
    Dim ArchivoNoEncontrado As String
    Dim Descripcion As String
    
    Dim rs As New ADODB.Recordset
    Dim Strlotes As String
    
    Set rsGrilla = New ADODB.Recordset
    Pasar = False
    
    Caja = InputBox(" INGRESE LA CAJA ")
    Descripcion = InputBox("Descripcion")
    Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " FROM  DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " Where DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = " & Caja & " AND  DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION = '" & Descripcion & "'"
    Sql = Sql & " AND DOCUMENTOS_DIGITALES.ESTADO = 'LISTA PARA EXPORTAR'"
    Sql = Sql & " GROUP BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
    Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
    Do While Not rs.EOF
        Strlotes = Strlotes & "," & rs!ID_DOCUMENTOS_DIGITALES_LOTE
        rs.MoveNext
    Loop
    
   If Strlotes = "" Then
    MsgBox "No existen lotes a exportar para esa caja "
    Exit Sub
   
   End If
   
   
    
    
    
        Sql = "  SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN,"
        Sql = Sql & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.ESTADO, DOCUMENTOS_DIGITALES.CONVERSION_TIPO_ARCHIVO,"
        Sql = Sql & " DOCUMENTOS_DIGITALES.DIRECTORIO_PASO ,  DOCUMENTOS_DIGITALES.CANTIDAD_IMAGENES ,  DOCUMENTOS_DIGITALES.ID  AS IDIMAGEN"
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & " DOCUMENTOS_DIGITALES ON"
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & " Where ID_DOCUMENTOS_DIGITALES_LOTE in(" & Mid(Strlotes, 2) & ")"
        Sql = Sql & " AND ESTADO = 'LISTA PARA EXPORTAR'"
        Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
        
    Rem RS.CursorLocation = adUseClient
    
   Set rs = New ADODB.Recordset
   
       rs.Open Sql, strConBasa
  
    If rs.EOF Then
        MsgBox "nO EXISTE LA CAJA"
        Exit Sub
        
    End If
    
        
        
        LoteNombreArchivoTxt = "LOTE" & Caja & "_" & Trim(Descripcion)
        PasoImagenes = "\\222.15.19.251\ImagenesPDF\"
       Rem  PasoExp = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS" & "\" & LoteNombreArchivoTxt
        PasoExp = "D:\EXPORTADAS" & "\" & LoteNombreArchivoTxt
       
    If Dir(PasoExp, vbDirectory) = "" Then
        FileSystem.MkDir PasoExp
        PasoExpRaiz = PasoExp
        PasoExpImagenes = PasoExpRaiz & "\IMAGENES"
        FileSystem.MkDir PasoExpImagenes
     Else
        MsgBox "Los directorios ya existen"
        Exit Sub
    End If
    
    
    
  Rem   I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM
    NombreArchivoTxt = PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".TXT"
    Open NombreArchivoTxt For Append As #1
   Rem     Encabezado = Chr(34) & "DocumentFileName" & Chr(34) & "," & Chr(34) & "PageCount" & Chr(34) & "," & Chr(34) & "Guia" & Chr(34) & "," & Chr(34) & "Fecha" & Chr(34) & "," & Chr(34) & "Sucursal" & Chr(34) & "," & Chr(34) & "Caja" & Chr(34) & "," & Chr(34) & "Lote" & Chr(34)
    Encabezado = Chr(34) & "DocumentFileName" & Chr(34) & "," & Chr(34) & "PageCount" & Chr(34) & "," & Chr(34) & "Guia" & Chr(34) & "," & Chr(34) & "Lote" & Chr(34) & "," & Chr(34) & "Caja" & Chr(34) & "," & Chr(34) & "Sucursal" & Chr(34) & "," & Chr(34) & "Fecha" & Chr(34)
    Print #1, Encabezado
        Do While Not rs.EOF
            CantImagenes = CantImagenes + 1
            
            Sql = " Update DOCUMENTOS_DIGITALES"
            Sql = Sql & " SET ESTADO ='EXPORTADA'"
            Sql = Sql & " Where ID = " & rs!IDIMAGEN
            ExecutarSql Sql
            
            NombreArchivoImagenPDF = Format(rs!FK_CAJAS, "00000000") & "_" & Format(rs!ID_DOCUMENTOS_DIGITALES_LOTE, "000000") & "_" & Format(rs!IMAGEN_ORIGEN, "0000") & "_" & rs!IDIMAGEN & ".pdf"
         Rem  NombreArchivoImagenPDF = RS!IDIMAGEN & ".pdf"
'            NombreArchivoImagenPDF = Mid(rsGrilla!BatchPgDta, 1, 16) & ".pdf"
            NombreArchivoImagenTIF = Format(rs!FK_CAJAS, "00000000") & "_" & Format(rs!ID_DOCUMENTOS_DIGITALES_LOTE, "000000") & "_" & Format(rs!IMAGEN_ORIGEN, "0000") & "_" & rs!IDIMAGEN & ".pdf"
            If IsNull(rs!Cantidad_Imagenes) Then
                CANTIDAD_IMAGEN = 1
            Else
                CANTIDAD_IMAGEN = rs!Cantidad_Imagenes
            End If
            Rem  Datos = Chr(34) & "//" & LoteNombreArchivoTxt & "/IMAGENES/" & NombreArchivoImagenPDF & Chr(34) & "," & Chr(34) & CANTIDAD_IMAGEN & Chr(34) & "," & Chr(34) & RS!NRO_DESDE & Chr(34) & "," & Chr(34) & Format(Now, "DD/MM/YYYY") & Chr(34) & "," & Chr(34) & "MENDOZA" & Chr(34) & "," & Chr(34) & Caja & Chr(34) & "," & Chr(34) & LoteNombreArchivoTxt & Chr(34)
            Rem Chr (34) & "//" & LoteNombreArchivoTxt & "/IMAGENES/" & NombreArchivoImagenPDF & Chr(34) & "," & Chr(34) & CANTIDAD_IMAGEN & Chr(34) & "," & Chr(34) & RS!NRO_DESDE & Chr(34) & "," & Chr(34) & LoteNombreArchivoTxt & Chr(34) & "," & Chr(34) & Caja & Chr(34) & "," & Chr(34) & "MENDOZA" & Chr(34) & "," & Chr(34) & Format(Now, "DD/MM/YYYY") & Chr(34)
            Datos = Chr(34) & "//" & LoteNombreArchivoTxt & "/IMAGENES/" & NombreArchivoImagenPDF & Chr(34) & "," & Chr(34) & CANTIDAD_IMAGEN & Chr(34) & "," & Chr(34) & rs!NRO_DESDE & Chr(34) & "," & Chr(34) & LoteNombreArchivoTxt & Chr(34) & "," & Chr(34) & Caja & Chr(34) & "," & Chr(34) & "MENDOZA" & Chr(34) & "," & Chr(34) & Format(Now, "DD/MM/YYYY") & Chr(34)
            Print #1, Datos
             
             
            Rem  ESTA MAL FileCopy "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3), PasoExpImagenes & "\" & NombreArchivoImagenTIF
             
             If Dir(PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".PDF", vbArchive) <> "" Then
                 FileCopy PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".PDF", PasoExpImagenes & "\" & NombreArchivoImagenPDF
                Else
                
                    If Dir("\\222.15.19.251\Imagenes" & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".TIF", vbArchive) <> "" Then
                     FileCopy "\\222.15.19.251\Imagenes" & "\" & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".TIF", PasoExpImagenes & "\" & NombreArchivoImagenTIF
                    Else
                      MsgBox "ERROR"
                    End If
                
            End If
            
            rs.MoveNext
            Debug.Print CantImagenes
        Loop
    Close #1
    ControlXML = "<Batch>" & vbCr
    ControlXML = ControlXML & " <Statistics>" & vbCr
    ControlXML = ControlXML & "     <DocumentCount>" & CantImagenes & "</DocumentCount>" & vbCr
    ControlXML = ControlXML & " </Statistics>" & vbCr
    ControlXML = ControlXML & "</Batch>" & vbCr
    Open PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".XML" For Append As #1
        Print #1, ControlXML
    Close #1
    
    If ArchivoNoEncontrado <> "" Then
                MsgBox (" NO se enconcotro el archivo")
                
                Clipboard.Clear
               Clipboard.SetText ArchivoNoEncontrado
                
    End If
    
    Sql = " Update DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " SET LOTE_ESTADO ='EXPORTADO'"
    Sql = Sql & " , FECHA_EXPORTACION=" & SysDateMinutoSegundo
    Sql = Sql & "  Where ID_DOCUMENTOS_DIGITALES_LOTE IN( " & Mid(Strlotes, 2) & ")"
    
    ExecutarSql Sql

    
    MsgBox "terminados"
End Sub

Private Sub Command2_Click()
    Dim Sql As String
    Dim Caja As Long
    Dim Encabezado As String
    Dim NombreArchivoImagenPDF As String
    Dim NombreArchivoImagenTIF As String
    Dim NombreArchivoTxt As String
    Dim LoteNombreArchivoTxt As String
    Dim PasoInicio As String
    Dim PasoImagenes As String
    Dim PasoExpImagenes  As String
    Dim PasoExp As String
    Dim Datos As String
    Dim CantImagenes As Long
    Dim ControlXML As String
    Dim PasoExpRaiz As String
    Dim CANTIDAD_IMAGEN As Integer
    Dim ArchivoNoEncontrado As String
    
    Set rsGrilla = New ADODB.Recordset
    Pasar = False
    Sql = " SELECT  ID, Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, NRO_DESDE2, BatchPgDta  , CANTIDAD_IMAGEN"
    Sql = Sql & " From TELEFORM_DIGITAL"
    Sql = Sql & " WHERE     (BatchNo = " & InputBox("Ingrese el lote") & ") and NRO_DESDE2 > 100 "
    Sql = Sql & " ORDER BY ID DESC"
    
    
    
    Set rsGrilla = New ADODB.Recordset
        rsGrilla.Open Sql, strConBasa
        Caja = Mid(rsGrilla!BatchPgDta, 1, 7)
        LoteNombreArchivoTxt = "LOTE" & Caja
        PasoImagenes = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM" & "\" & Caja
        PasoExp = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS" & "\" & "LOTE" & Caja
    If Dir(PasoExp, vbDirectory) = "" Then
        FileSystem.MkDir PasoExp
        PasoExpRaiz = PasoExp
        PasoExpImagenes = PasoExpRaiz & "\IMAGENES"
        FileSystem.MkDir PasoExpImagenes
     Else
        MsgBox "Los directorios ya existen"
        Exit Sub
    End If
  Rem   I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM
    NombreArchivoTxt = PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".TXT"
    Open NombreArchivoTxt For Append As #1
    Encabezado = Chr(34) & "DocumentFileName" & Chr(34) & "," & Chr(34) & "PageCount" & Chr(34) & "," & Chr(34) & "Guia" & Chr(34) & "," & Chr(34) & "Fecha" & Chr(34) & "," & Chr(34) & "Sucursal" & Chr(34) & "," & Chr(34) & "Caja" & Chr(34) & "," & Chr(34) & "Lote" & Chr(34)
    Print #1, Encabezado
        Do While Not rsGrilla.EOF
            CantImagenes = CantImagenes + 1
            NombreArchivoImagenPDF = Mid(rsGrilla!BatchPgDta, 1, 13) & ".pdf"
            NombreArchivoImagenTIF = Mid(rsGrilla!BatchPgDta, 1, 13) & ".TIF"
'            NombreArchivoImagenPDF = Mid(rsGrilla!BatchPgDta, 1, 16) & ".pdf"
'            NombreArchivoImagenTIF = Mid(rsGrilla!BatchPgDta, 1, 16) & ".TIF"
            If IsNull(rsGrilla!CANTIDAD_IMAGEN) Then
                CANTIDAD_IMAGEN = 1
            Else
                CANTIDAD_IMAGEN = rsGrilla!CANTIDAD_IMAGEN
            End If
            Datos = Chr(34) & "//" & LoteNombreArchivoTxt & "/IMAGENES/" & NombreArchivoImagenPDF & Chr(34) & "," & Chr(34) & CANTIDAD_IMAGEN & Chr(34) & "," & Chr(34) & rsGrilla!NRO_DESDE2 & Chr(34) & "," & Chr(34) & Format(Now, "DD/MM/YYYY") & Chr(34) & "," & Chr(34) & "MENDOZA" & Chr(34) & "," & Chr(34) & Caja & Chr(34) & "," & Chr(34) & LoteNombreArchivoTxt & Chr(34)
            Print #1, Datos
             
             
            Rem  ESTA MAL FileCopy "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3), PasoExpImagenes & "\" & NombreArchivoImagenTIF
             
             If Dir(PasoImagenes & "\" & NombreArchivoImagenTIF, vbArchive) <> "" Then
                FileCopy PasoImagenes & "\" & NombreArchivoImagenTIF, PasoExpImagenes & "\" & NombreArchivoImagenTIF
            Else
                 
                ArchivoNoEncontrado = ArchivoNoEncontrado & vbCrLf & "NO esta: " & PasoImagenes & "\" & NombreArchivoImagenTIF & "   Sacar:  " & "\\PCTELEMEMO1" & rsGrilla!Suspense_File
                FileCopy "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3), PasoExpImagenes & "\" & NombreArchivoImagenTIF
            End If
            
            rsGrilla.MoveNext
            Debug.Print CantImagenes
        Loop
    Close #1
    ControlXML = "<Batch>" & vbCr
    ControlXML = ControlXML & " <Statistics>" & vbCr
    ControlXML = ControlXML & "     <DocumentCount>" & CantImagenes & "</DocumentCount>" & vbCr
    ControlXML = ControlXML & " </Statistics>" & vbCr
    ControlXML = ControlXML & "</Batch>" & vbCr
    Open PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".XML" For Append As #1
        Print #1, ControlXML
    Close #1
    
    If ArchivoNoEncontrado <> "" Then
                MsgBox (" NO se enconcotro el archivo")
                
                Clipboard.Clear
               Clipboard.SetText ArchivoNoEncontrado
                
    End If
    MsgBox "terminados"
End Sub



Private Sub Command3_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Nombre As String




Sql = "SELECT     BatchNo, SUBSTRING(CONVERT(char, NombreArchivoNumero), 1, 7) AS CAJA, ID, NombreArchivoNumero"
Sql = Sql & " From basasql.dbo.TELEFORM_BARRA"
Sql = Sql & " WHERE     (SUBSTRING(CONVERT(char, NombreArchivoNumero), 1, 7) = '1039571')"

rs.Open Sql, strConBasa
Do While Not rs.EOF

Nombre = 1039517 & Mid(rs!NombreArchivoNumero, 8)

Sql = " Update TELEFORM_BARRA"
Sql = Sql & " Set NombreArchivoNumero = " & Nombre
Sql = Sql & " Where ID = " & Nombre
ExecutarSql Sql

    rs.MoveNext
Loop



End Sub

Private Sub Command5_Click()
Dim Sql As String
Dim lote As String
lote = InputBox("Ingrese el lote")
Dim LargoBarra As Integer

Sql = "  SELECT id, BatchNo, BARRA"
Sql = Sql & " From TELEFORM_BARRA"
Sql = Sql & "  Where BatchNo = " & lote
Sql = Sql & "  ORDER BY BARRA"

Dim rs2 As New ADODB.Recordset
Dim ID As String

rs2.Open Sql, strConBasa


        Do While Not rs2.EOF
            LargoBarra = Len(Trim(rs2!BARRA))
            Sql = " Update TELEFORM_BARRA "
            Sql = Sql & " SET LARGO_BARRA =" & LargoBarra
            Sql = Sql & " Where ID = " & rs2!ID
            ExecutarSql Sql
            rs2.MoveNext
        Loop
        
        
        Dim Rslargo As New ADODB.Recordset
        Dim Msg As String
        
        
       
        Sql = " SELECT LEN(BARRA) AS Largo, COUNT(*) as Cant"
        Sql = Sql & " From TELEFORM_BARRA"
        Sql = Sql & " Where BatchNo = " & lote
        Sql = Sql & " GROUP BY LEN(BARRA)"
        Set Rslargo = New ADODB.Recordset
        Rslargo.Open Sql, strConBasa
        
      Do While Not Rslargo.EOF
        Msg = "El largo es : " & Rslargo!largo & " La Cantidad es :" & Rslargo!cant & vbCrLf & Msg
        Rslargo.MoveNext
      Loop
      
        MsgBox Msg
        
        Sql = "  SELECT ID, Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, BARRA,  BatchPgDta "
        Sql = Sql & " From TELEFORM_BARRA"
        Sql = Sql & " Where BatchNo = " & lote
        Sql = Sql & " AND LARGO_BARRA in(" & InputBox("Ingrese el largo para analizar , puede ser 1,5 ,6 ") & ")"
        Sql = Sql & "  ORDER BY BARRA"
        
        
        Set rsGrilla = New ADODB.Recordset
        rsGrilla.CursorLocation = adUseClient
        rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic

        Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
        grdIndexarImagenes.ReBind
        grdIndexarImagenes.Refresh


End Sub

Private Sub Command6_Click()
    Dim Sql As String
    Dim Caja As Long
    Dim Encabezado As String
    Dim NombreArchivoImagenPDF As String
    Dim NombreArchivoImagenTIF As String
    Dim NombreArchivoTxt As String
    Dim LoteNombreArchivoTxt As String
    Dim PasoInicio As String
    Dim PasoImagenes As String
    Dim PasoExpImagenes  As String
    Dim PasoExp As String
    Dim Datos As String
    Dim CantImagenes As Long
    Dim ControlXML As String
    Dim PasoExpRaiz As String
    Dim CANTIDAD_IMAGEN As Integer
    
    Set rsGrilla = New ADODB.Recordset
    Pasar = False
    Sql = " SELECT  ID, Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, NRO_DESDE2, BatchPgDta  , CANTIDAD_IMAGEN"
    Sql = Sql & " From TELEFORM_DIGITAL"
    Sql = Sql & " WHERE     (BatchNo = " & InputBox("Ingrese el lote") & ") and NRO_DESDE2 > 100 "
    Sql = Sql & " ORDER BY ID DESC"
    
    Sql = "  SELECT     Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, NRO_DESDE2, BatchPgDta, CANTIDAD_IMAGEN"
Sql = Sql & " From TELEFORM_DIGITAL"
Sql = Sql & " GROUP BY Suspense_File, CSID, Verify_Wks, Form_Id, BatchNo, NRO_DESDE2, BatchPgDta, CANTIDAD_IMAGEN"
Sql = Sql & " Having (BatchNo = 1152) And (NRO_DESDE2 > 100)"
    
    Set rsGrilla = New ADODB.Recordset
    rsGrilla.Open Sql, strConBasa
    Caja = Mid(rsGrilla!BatchPgDta, 1, 7)
    LoteNombreArchivoTxt = "LOTE" & Caja
    PasoImagenes = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM" & "\" & Caja
    PasoExp = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\EXPORTADAS" & "\" & "LOTE" & Caja
    
  
  
   
        Do While Not rsGrilla.EOF
            
            NombreArchivoImagenTIF = Mid(rsGrilla!BatchPgDta, 1, 13) & ".TIF"
            If Dir(PasoImagenes & "\" & NombreArchivoImagenTIF) = "" Then
                MsgBox "NO esta " & NombreArchivoImagenTIF
            End If
            rsGrilla.MoveNext
        Loop
    

    MsgBox "terminados"
End Sub

Private Sub Command7_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = " SELECT     CANTIDAD_IMAGEN, BatchNo, CONVERT(numeric, SUBSTRING(BatchPgDta, 1, 8)) AS CAJA, NRO_DESDE2, Form_Id, BatchPgDta"
Sql = Sql & " From TELEFORM_DIGITAL"
Sql = Sql & "  WHERE     (Form_Id IN (53984, 22258)) AND (NOT (BatchPgDta IS NULL))"


Sql = "  SELECT     CANTIDAD_IMAGEN, BatchNo, CONVERT(numeric, SUBSTRING(BatchPgDta, 1, 7)) AS CAJA, NRO_DESDE2, Form_Id, BatchPgDta, ID, CAJABASA"
Sql = Sql & " From TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (Form_Id IN (53984, 22258)) AND (NOT (BatchPgDta IS NULL))"
Sql = Sql & " ORDER BY CAJA"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    Sql = " Update TELEFORM_DIGITAL"
Sql = Sql & " Set CAJABASA = " & rs!Caja
Sql = Sql & " Where ID = " & rs!ID
ExecutarSql Sql
    rs.MoveNext
Loop



End Sub



Private Sub Command9_Click()
    Dim StrConAsp As String
    StrConAsp = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=False;User ID=sa;Initial Catalog=basa;Data Source=181.164.215.224"
    Dim con As New ADODB.Connection
    con.Open StrConAsp
    Dim rs As New ADODB.Recordset
    
Dim Sql As String

Sql = "SELECT     id, codigo, estado"
Sql = Sql & " From Elementos"
Sql = Sql & "  Where (clienteAsp_id = 1)  AND (clienteEmp_id = 1) "

    
    
     Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open Sql, StrConAsp, adOpenKeyset, adLockOptimistic

        Set grdIndexarImagenes.DataSource = rs.DataSource
        grdIndexarImagenes.ReBind
        grdIndexarImagenes.Refresh
    
    

End Sub

Private Sub Form_Activate()
Pasar = False
End Sub

Private Sub Form_Load()
Pasar = False
End Sub

Private Sub grdIndexarImagenes_Change()
'Dim Paso As String
'    On Error GoTo salir
'           Paso = "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3)
'
'            If Dir(Paso) <> "" Then
'                ctlVerImagenes1.PonerImagen Paso
'            Else
'                MsgBox "No existe la imagen"
'            End If
'   Exit Sub
'salir:
'MsgBox "No existe la imagen"

End Sub

Private Sub grdIndexarImagenes_Click()
PonerImagenAndesmar "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3)
End Sub

Private Sub grdIndexarImagenes_DblClick()
    Dim Paso
    Paso = "\\PCTELEMEMO1" & Mid(rsGrilla!Suspense_File, 3)
    Clipboard.Clear
    Clipboard.SetText Paso
    MsgBox "Paso Copiado"
End Sub

Private Sub grdIndexarImagenes_GotFocus()
    Pasar = True
End Sub

Private Sub grdIndexarImagenes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 

  If Pasar = True Then
    grdIndexarImagenes.Col = 1
    PonerImagenAndesmar "\\PCTELEMEMO1" & Mid(grdIndexarImagenes.Text, 3)
    grdIndexarImagenes.Col = 6
  End If
End Sub

Public Sub PonerImagenAndesmar(Paso As String)

    On Error GoTo salir
           
            
            If Dir(Paso) <> "" Then
                ctlVerImagenes1.PonerImagen Paso
            Else
                MsgBox "No existe la imagen"
            End If
   Exit Sub
salir:
MsgBox "No existe la imagen"
End Sub

