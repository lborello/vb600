VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControlHorario 
   Caption         =   "Control Horario"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   10155
   Begin VB.CommandButton cmdImportarDatos 
      Caption         =   "Importar Datos"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Actualizar Fecha"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Excel Todos"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Actualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid grdHorario 
      Height          =   4935
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton Command3 
      Caption         =   "Excel"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sale"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingreso"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtHorarioIngreso 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtUsuario 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Horario"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmControlHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim rsActualizar As ADODB.Recordset



Private Sub Actualizar_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

If InputBox("Ingrese la clave") <> "181177" Then
 MsgBox "Solo Administradores"
Exit Sub

End If

On Error GoTo salir:

Sql = " SELECT     CONTROLHORARIOS.ID, PERSONAL.IDPERSONAL, PERSONAL.APELLIDO, PERSONAL.NOMBRE, CONTROLHORARIOS.FECHA,"
Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_1, CONTROLHORARIOS.HORA_SALIDA_1, CONTROLHORARIOS.SUMA_1,"
Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_2, CONTROLHORARIOS.HORA_SALIDA_2, CONTROLHORARIOS.SUMA_2,"
Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_3, CONTROLHORARIOS.HORA_SALIDA_3, CONTROLHORARIOS.SUMA_3,"
Sql = Sql & " CONTROLHORARIOS.TOTAL_HORA"
Sql = Sql & " FROM PERSONAL INNER JOIN"
Sql = Sql & " CONTROLHORARIOS ON PERSONAL.IDPERSONAL = CONTROLHORARIOS.FK_PERSONAL"
Sql = Sql & " Where Personal.IDPERSONAL = " & txtUsuario.Text

    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly

    Set grdHorario.DataSource = rs.DataSource
       grdHorario.DataMember = rs.DataMember


Exit Sub

salir:
MsgBox "Ingrese el personal"
MsgBox Err.Description
End Sub

Private Sub cmdImportarDatos_Click()
Dim DATO As String
   
    
    Dim F As Integer
    Dim Lineas As Long
    Dim str_Linea As String
    Dim FK_PERSONAL As Integer
    Dim fecha As String
    Dim Hora As String
    Dim Minuto As String
    Dim FECHA_HORA As String
Dim Sql As String
    ' Número de archivo libre
    F = FreeFile

    ' Abre el archivo de texto
    Open "C:\Horario\actual.txt" For Input As #F

    'Recorre todo el archivo de texto _
     linea por linea hasta el final
    Do
        'Lee una línea
        Line Input #F, str_Linea
 Rem        MsgBox str_Linea
        
        If Trim(str_Linea) <> "" Then
        If CInt(Mid(str_Linea, 1, 10)) > 1000 Then
            FK_PERSONAL = CInt(Mid(str_Linea, 1, 10)) - 1000
        Else
            FK_PERSONAL = CInt(Mid(str_Linea, 1, 10))
        End If
        fecha = Mid(str_Linea, 19, 2) & "/" & Mid(str_Linea, 16, 2) & "/" & Mid(str_Linea, 11, 4)
        Hora = Mid(str_Linea, 22, 2)
        Minuto = Mid(str_Linea, 25, 2)
        
        FECHA_HORA = "'" & fecha & " " & Format(Hora, "00") & ":" & Format(Minuto, "00") & ":00'"
        
        
        fecha = "'" & fecha & "'"
        Sql = " INSERT Into TEM_HORARIOS("
        Sql = Sql & " FK_PERSONAL"
        Sql = Sql & " , Fecha"
        Sql = Sql & " , Hora"
        Sql = Sql & " , MINUTO"
        Sql = Sql & " , FECHA_HORA )"
        Sql = Sql & " VALUES  ("
        Sql = Sql & FK_PERSONAL
        Sql = Sql & " ," & fecha
        Sql = Sql & " , " & Hora
        Sql = Sql & " , " & Minuto
        Sql = Sql & " , " & FECHA_HORA
        Sql = Sql & ")"
        ExecutarSql Sql
        
        
        ' Incrementa la cantidad de lineas leidas
        Lineas = Lineas + 1
        End If
        
    
    ' Leerá hasta que llegue al fin de archivo
    Loop While Not EOF(F)

    ' Cierra el archivo de texto abierto
    Close #F

    ' Retorna a la función el número de lineas del fichero



'Insert Top(1000)
'Into TEM_HORARIOS(FK_PERSONAL, fecha, HORA_INGESO_1, HORA_INGESO_2, HORA_SALIDA_1, HORA_SALIDA_2)
'VALUES        (45, CONVERT(DATETIME, '2006-10-10 00:00:00', 102), 11, 11, 11, 110)


End Sub

Private Sub Command1_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String


    Sql = " SELECT  FK_PERSONAL, FECHA, HORA_INGRESO_1, HORA_SALIDA_1, HORA_INGRESO_2, HORA_SALIDA_2, HORA_INGRESO_3, HORA_SALIDA_3 FROM CONTROLHORARIOS "
    Sql = Sql & " WHERE FK_PERSONAL = " & txtUsuario.Text
    Sql = Sql & " AND FECHA = " & FechaFormato(txtHorarioIngreso.Text)
    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic
 
    
If Not rs.EOF Then
    If IsNull(rs!HORA_INGRESO_1) Then
        rs!HORA_INGRESO_1 = txtHorarioIngreso.Text
    Else
        If IsNull(rs!HORA_INGRESO_2) Then
            rs!HORA_INGRESO_2 = txtHorarioIngreso.Text
        Else
            If IsNull(rs!HORA_INGRESO_3) Then
                rs!HORA_INGRESO_3 = txtHorarioIngreso.Text
            End If
        End If
    End If
    rs.Update
 Else
    rs.AddNew
    rs!FK_PERSONAL = txtUsuario.Text
    rs!fecha = Format(txtHorarioIngreso.Text, "DD/MM/YYYY")
    rs!HORA_INGRESO_1 = txtHorarioIngreso.Text
    rs.Update
 
 End If
 MsgBox "Terminado"

End Sub

Private Sub Command2_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String


    Sql = " SELECT  FK_PERSONAL, FECHA, HORA_INGRESO_1, HORA_SALIDA_1, HORA_INGRESO_2, HORA_SALIDA_2, HORA_INGRESO_3, HORA_SALIDA_3 FROM CONTROLHORARIOS "
    Sql = Sql & " WHERE FK_PERSONAL = " & txtUsuario.Text
    Sql = Sql & " AND FECHA = " & FechaFormato(txtHorarioIngreso.Text)
    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic
 
    
If Not rs.EOF Then
    If IsNull(rs!HORA_SALIDA_1) Then
        rs!HORA_SALIDA_1 = txtHorarioIngreso.Text
    Else
        If IsNull(rs!HORA_SALIDA_2) Then
            rs!HORA_SALIDA_2 = txtHorarioIngreso.Text
        Else
            If IsNull(rs!HORA_SALIDA_3) Then
                rs!HORA_SALIDA_3 = txtHorarioIngreso.Text
            End If
        End If
    End If
    rs.Update
 Else
    rs.AddNew
    rs!FK_PERSONAL = txtUsuario.Text
    rs!fecha = Format(txtHorarioIngreso.Text, "DD/MM/YYYY")
    rs!HORA_SALIDA_1 = txtHorarioIngreso.Text
    rs.Update
 
 End If

 MsgBox "Terminado"

End Sub

Private Sub Command3_Click()
CopiarDatosGrilla grdHorario
End Sub

Private Sub Command4_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim SUMA_1 As Double
    Dim SUMA_2 As Double
    Dim SUMA_3 As Double
    Dim TOTAL_HORA As Double

Sql = " SELECT     ID, HORA_INGRESO_1, HORA_SALIDA_1, SUMA_1, HORA_INGRESO_2, HORA_SALIDA_2, SUMA_2, HORA_INGRESO_3, HORA_SALIDA_3, SUMA_3,"
Sql = Sql & " TOTAL_HORA"
Sql = Sql & " From CONTROLHORARIOS"
Sql = Sql & "  ORDER BY ID"
rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic

Do While Not rs.EOF
    If Not IsNull(rs!HORA_INGRESO_1) Then
        If Not IsNull(rs!HORA_SALIDA_1) Then
                SUMA_1 = DateDiff("s", rs!HORA_INGRESO_1, rs!HORA_SALIDA_1)
                rs!SUMA_1 = Replace(Format(SUMA_1 / 60 / 60, "00.00"), ",", ".")
        End If
    
    End If
    
    If Not IsNull(rs!HORA_INGRESO_2) Then
        If Not IsNull(rs!HORA_SALIDA_2) Then
                SUMA_2 = DateDiff("s", rs!HORA_INGRESO_2, rs!HORA_SALIDA_2)
            rs!SUMA_2 = Replace(Format(SUMA_2 / 60 / 60, "00.00"), ",", ".")
        End If
    
    End If
    If Not IsNull(rs!HORA_INGRESO_3) Then
        If Not IsNull(rs!HORA_SALIDA_3) Then
                SUMA_3 = DateDiff("s", rs!HORA_INGRESO_3, rs!HORA_SALIDA_3)
            rs!SUMA_3 = Replace(Format(SUMA_3 / 60 / 60, "00.00"), ",", ".")
        
        
        End If
    
    End If
    
    rs!TOTAL_HORA = Replace(Format(SUMA_1 / 60 / 60 + SUMA_2 / 60 / 60 + SUMA_3 / 60 / 60, "00.00"), ",", ".")
    rs.Update
    rs.MoveNext
Loop


End Sub

Private Sub Command5_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset




On Error GoTo salir:

Sql = " SELECT     CONTROLHORARIOS.ID, PERSONAL.IDPERSONAL, PERSONAL.APELLIDO, PERSONAL.NOMBRE, CONTROLHORARIOS.FECHA,"
Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_1, CONTROLHORARIOS.HORA_SALIDA_1, CONTROLHORARIOS.SUMA_1,"
Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_2, CONTROLHORARIOS.HORA_SALIDA_2, CONTROLHORARIOS.SUMA_2,"
Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_3, CONTROLHORARIOS.HORA_SALIDA_3, CONTROLHORARIOS.SUMA_3,"
Sql = Sql & " CONTROLHORARIOS.TOTAL_HORA"
Sql = Sql & " FROM PERSONAL INNER JOIN"
Sql = Sql & " CONTROLHORARIOS ON PERSONAL.IDPERSONAL = CONTROLHORARIOS.FK_PERSONAL"
Sql = Sql & " WHERE CONTROLHORARIOS.FECHA > " & FechaFormato(InputBox("INGRESE LA FECHA "))
Sql = Sql & " ORDER BY FK_PERSONAL, FECHA"

    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly

    Set grdHorario.DataSource = rs.DataSource
       grdHorario.DataMember = rs.DataMember


Exit Sub

salir:
MsgBox Err.Description
End Sub

Private Sub Command6_Click()
Dim Sql As String
If InputBox("Ingrese la Clave") = "2338" Then

    If MDIfrmInicio.StaInicio.Panels(2).Text <> 47 Then
     MsgBox "Solo Administradores"
    Exit Sub
    
    End If
        On Error GoTo salir:
        Sql = " SELECT    FK_PERSONAL,  CONTROLHORARIOS.FECHA,"
        Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_1, CONTROLHORARIOS.HORA_SALIDA_1, CONTROLHORARIOS.SUMA_1,"
        Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_2, CONTROLHORARIOS.HORA_SALIDA_2, CONTROLHORARIOS.SUMA_2,"
        Sql = Sql & " CONTROLHORARIOS.HORA_INGRESO_3, CONTROLHORARIOS.HORA_SALIDA_3, CONTROLHORARIOS.SUMA_3,"
        Sql = Sql & " CONTROLHORARIOS.TOTAL_HORA"
        Sql = Sql & " FROM  CONTROLHORARIOS "
        Sql = Sql & " Where FECHA = '" & InputBox("Ingrese Fecha", "Fecha", Format(Now, "dd/mm/yyyy")) & "'"
        Sql = Sql & "order by FK_PERSONAL "
        Set rsActualizar = New ADODB.Recordset
        rsActualizar.CursorLocation = adUseClient
        rsActualizar.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic
        Set grdHorario.DataSource = rsActualizar.DataSource
        grdHorario.DataMember = rsActualizar.DataMember
    
End If
Exit Sub
salir:
MsgBox "Ingrese el personal"
MsgBox Err.Description
End Sub

Private Sub Command7_Click()
Dim rsPersonal As New ADODB.Recordset
Dim rsDia As New ADODB.Recordset
Dim RsControl As New ADODB.Recordset
Dim Hora_Ingreso As String
Dim Hora_Salida As String
Dim Diferencia_horas As Double
Dim Diferencia_Minutos_total As Integer
Dim Diferencia_Minutos  As Integer
Dim Sql As String

Dim C_FK_PERSONAL As String
Dim C_FECHA  As String
Dim C_HORA_INGRESO_1 As String
Dim C_HORA_SALIDA_1 As String
Dim C_SUMA_1 As String
Dim C_HORA As String
Dim V_HORA As String
Dim C_MINUTO As String
Dim V_MINUTO As String
Dim C_TIPO_DIA As String
Dim C_NOMBRE_DIA As String




Sql = " SELECT       IDPERSONAL, NOMBRE, APELLIDO, HORA_INGRESO_INICIO, HORA_INGRESO_FIN, HORA_SALIDA_INICIO, HORA_SALIDA_FIN"
Sql = Sql & " From Personal"
Sql = Sql & " Where (HORA_INGRESO_INICIO > 0) "


rsPersonal.Open Sql, strConBasa

Do While Not rsPersonal.EOF
    Sql = " SELECT  ID, TIPO_DIA, DIA , NOMBRE_DIA, CONTROL , CONVERT (char,  DIA ,103 ) AS DIA_FORMATO "
    Sql = Sql & "  From Dia "
    Sql = Sql & "  WHERE DIA >  '29/06/2016'"
    Rem Sql = Sql & "  WHERE        (DIA BETWEEN CONVERT(DATETIME, '2016-07-01 00:00:00', 102) AND CONVERT(DATETIME, '2016-07-31 00:00:00', 102))"
    Sql = Sql & "  ORDER BY DIA"
    Set rsDia = New ADODB.Recordset
      rsDia.Open Sql, strConBasa
      Do While Not rsDia.EOF
        C_TIPO_DIA = "'" & rsDia!TIPO_DIA & "'"
        C_NOMBRE_DIA = "'" & UCase(Trim(rsDia!NOMBRE_DIA)) & "'"
        Sql = " SELECT min(FECHA_HORA) AS FECHA_HORA_INICIO"
        Sql = Sql & " From TEM_HORARIOS"
        Sql = Sql & " WHERE "
        Sql = Sql & " FECHA = '" & rsDia!Dia & "'"
        Sql = Sql & " AND FK_PERSONAL =" & rsPersonal!idPersonal
        
        Set RsControl = New ADODB.Recordset
        RsControl.Open Sql, strConBasa
        
        If Not RsControl.EOF Then
         If Not IsNull(RsControl!FECHA_HORA_INICIO) Then
               Hora_Ingreso = RsControl!FECHA_HORA_INICIO
            Else
                Hora_Ingreso = ""
            End If
            
        Else
        Hora_Ingreso = ""
        End If

If Hora_Ingreso <> "" Then
        Sql = " SELECT MAX(FECHA_HORA) AS FECHA_HORA_FIN"
        Sql = Sql & "  From TEM_HORARIOS"
        Sql = Sql & "  WHERE "
        Sql = Sql & " FECHA = '" & rsDia!Dia & "'"
        Sql = Sql & "  AND FK_PERSONAL =" & rsPersonal!idPersonal
        
        Set RsControl = New ADODB.Recordset
        RsControl.Open Sql, strConBasa
        
        If Not RsControl.EOF Then
            Hora_Salida = RsControl!FECHA_HORA_FIN
        Else
            Hora_Salida = ""
        End If

        If Hora_Ingreso = Hora_Salida Then
                C_HORA_INGRESO_1 = "'" & Hora_Ingreso & "'"
                C_HORA_SALIDA_1 = "'" & Hora_Salida & "'"
                C_SUMA_1 = "0"
        
        Else
                Diferencia_Minutos_total = DateDiff("n", Hora_Ingreso, Hora_Salida)
                If Diferencia_Minutos_total > 60 Then
                    Diferencia_horas = Format(CSng(Diferencia_Minutos_total / 60), "##.#")
                    V_HORA = Mid(Format(CStr((Diferencia_Minutos_total / 60)), "00.00"), 1, 2)
                    V_MINUTO = CInt(CInt(Mid(Format(CStr((Diferencia_Minutos_total / 60)), "00.00"), 4, 2)) * 0.6)
        
                Else
                    C_SUMA_1 = "'" & Trim(rsDia!DIA_FORMATO) & "'"
                End If
                
                C_HORA_INGRESO_1 = "'" & Hora_Ingreso & "'"
                C_HORA_SALIDA_1 = "'" & Hora_Salida & "'"
                C_SUMA_1 = "'" & Trim(rsDia!DIA_FORMATO) & " " & V_HORA & ":" & V_MINUTO & "'"
         End If
 
 Else
    C_HORA_INGRESO_1 = "NULL"
    C_HORA_SALIDA_1 = "NULL"
    C_SUMA_1 = "'" & Trim(rsDia!DIA_FORMATO) & "'"
 End If

        Sql = " INSERT INTO CONTROLHORARIOS"
        Sql = Sql & " ("
        Sql = Sql & " FK_PERSONAL"
        Sql = Sql & " , FECHA"
        Sql = Sql & " , HORA_INGRESO_1"
        Sql = Sql & " , HORA_SALIDA_1"
        Sql = Sql & " , SUMA_1"
        Sql = Sql & " , TIPO_DIA"
        Sql = Sql & " , NOMBRE_DIA"
        Sql = Sql & " )"
        Sql = Sql & " VALUES     "
        Sql = Sql & " ("
        Sql = Sql & rsPersonal!idPersonal
        Sql = Sql & " , '" & Trim(rsDia!DIA_FORMATO) & "'"
        Sql = Sql & " , " & C_HORA_INGRESO_1
        Sql = Sql & " , " & C_HORA_SALIDA_1
        Sql = Sql & " , " & C_SUMA_1
        Sql = Sql & " , " & C_TIPO_DIA
        Sql = Sql & " , " & C_NOMBRE_DIA
        Sql = Sql & " )"
        
        
        ExecutarSql Sql
        
        rsDia.MoveNext
      
      
      
      Loop
      
        

    rsPersonal.MoveNext
Loop


End Sub

Private Sub Form_Load()

 txtHorarioIngreso.Enabled = True



End Sub

Private Sub txtHorarioIngreso_DblClick()
If InputBox("Ingrese la clave") = "181177" Then
 txtHorarioIngreso.Enabled = True
End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   On Error GoTo salir
    If KeyAscii = 13 Then
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Sql = "SELECT     IDPERSONAL, NOMBRE, APELLIDO"
        Sql = Sql & " From Personal  Where IDPERSONAL = " & txtUsuario.Text
        rs.Open Sql, ConActiva, 0, 1
        If Not rs.EOF Then
            Label2.Caption = Trim(rs!Nombre) & " " & Trim(rs!Apellido)
            txtHorarioIngreso.Text = Replace(SysDate2, "'", "")
        End If
        
        
    
    End If

Exit Sub
salir:

    MsgBox Err.Description

End Sub
