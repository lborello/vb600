Attribute VB_Name = "Module1"
Public Sub CopiarDatosGrilla(Grilla As DataGrid)
Dim C As Integer
Dim R As Integer
Dim RSDATOS As ADODB.Recordset
Dim DATO As String
Dim ColGrilla As Integer
Dim DatoPuro As String
Set RSDATOS = New ADODB.Recordset

Set RSDATOS.DataSource = Grilla.DataSource
 On Error GoTo salir
 

For C = 0 To RSDATOS.Fields.Count - 1


    DATO = DATO & RSDATOS.Fields(C).Name & vbTab
 Next
    DATO = DATO & vbCrLf
    Do While Not RSDATOS.EOF
        For C = 0 To RSDATOS.Fields.Count - 1
        
        
            If Not IsNull(RSDATOS.Fields.Item(C).Value) Then
                DatoPuro = Replace(RSDATOS.Fields.Item(C).Value, vbCr, "")
                DatoPuro = Replace(DatoPuro, vbTab, "")
                DatoPuro = Replace(DatoPuro, vbCrLf, "")
                DatoPuro = Replace(DatoPuro, Chr(10), "")
                DatoPuro = CStr(DatoPuro)
                DATO = DATO & DatoPuro & vbTab
            Else
                DATO = DATO & "" & vbTab
            End If
        Next
        RSDATOS.MoveNext
        DATO = DATO & vbCrLf
    Loop
 Clipboard.Clear
 Clipboard.SetText DATO
 MsgBox "LOS DATOS FUERON COPIADOS"
salir:
If Err.Number <> 0 Then
    MsgBox Err.Description
    Exit Sub
End If
 
End Sub

Public Function FechaFormato(fecha As Variant)
    If UCase(fecha) <> "NULL" Then
        FechaFormato = " CONVERT(DATETIME, '" & Format(fecha, "YYYY-MM-DD") & " 00:00:00', 102)"
    Else
        FechaFormato = "NULL"
    End If
End Function

