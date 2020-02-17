Attribute VB_Name = "Module1"
Option Explicit
Public UserName  As String
Public Password As String


Public Function CollecionASql(Col As Collection) As String
    Dim i As Integer
    Dim dato As String
    For i = 1 To Col.Count
        dato = dato & Col.Item(i) & ","
    Next
    CollecionASql = Mid(dato, 1, Len(dato) - 1)
    
End Function
Public Function Digito_Verificador(dato As String) As Integer
Dim i As Integer
Dim Sumar As Integer

    For i = 1 To Len(Trim(dato))
        Sumar = Sumar + Mid(dato, i, 1)
    Next
      Digito_Verificador = Sumar
      
End Function

Public Function SendMail(sEmailRecipient As String, sEmailSubject As String, sEmailBody As String, Optional sAttachment1 As String, Optional sAttachment2 As String, Optional sAttachment3 As String)

'-----Send an Email Message using Outlook 98-----

'Developers Note:  In References, the Microsoft Outlook 98 Object Model must be selected for this to work
'                  Actual file is MSOUTL85.OLB

Dim emailOutlookApp As Outlook.Application

Dim emailNameSpace As Outlook.Namespace
Dim emailFolder As Outlook.MAPIFolder
Dim emailItem As Outlook.MailItem
Dim EmailRecipient As Recipient

'-----Open Outlook in a background process and the Inbox Folder-----
Set emailOutlookApp = CreateObject("Outlook.Application")
Set emailNameSpace = emailOutlookApp.GetNamespace("MAPI")
Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderInbox)

'Enable the next line to actually see Outlook Open
'emailFolder.Display

'-----Create a new mail message, set the recipient, subject, and body-----
Set emailItem = emailOutlookApp.CreateItem(olMailItem)
Set EmailRecipient = emailItem.Recipients.Add(sEmailRecipient)
emailItem.Subject = sEmailSubject
emailItem.Body = sEmailBody
If sAttachment1 <> "" Then
    emailItem.Attachments.Add sAttachment1
End If
If sAttachment2 <> "" Then
    emailItem.Attachments.Add sAttachment2
End If
If sAttachment3 <> "" Then
    emailItem.Attachments.Add sAttachment3
End If


'-----Send the Email-----
emailItem.Save
emailItem.Send

'-----Close the Outlook Application------
Rem emailOutlookApp.Quit

'----Inform User of Success-----
Rem  MsgBox "El Correo fue enviado", vbInformation

'-----Clear out the memory space held by variables-----
'Usually unnecessary but a good practice

Set emailNameSpace = Nothing
Set emailFolder = Nothing
Set emailItem = Nothing
Set emailOutlookApp = Nothing

End Function
Public Sub Movimiento(Cod_REMITO As Long, Cod_Tipo_Remito As Long, Cod_Operacion_Remito As Long, _
    Cod_Tipo_Almacemamiento As Integer, elementos As Collection, Cod_cliente As Long, Fecha As String)
'    Dim i As Integer
'    Dim Ssql As String
'        For i = 0 To Elementos.Count - 1
'            Ssql = "INSERT INTO MOVIMIENTOS_ELEMENTOS "
'            Ssql = Ssql & vbCrLf & " (COD_REMITO, COD_TIPO, COD_OPERACION,"
'            Ssql = Ssql & vbCrLf & " COD_TIPO_ALMACENAMIENTO, ELEMENTO, COD_CLIENTE,"
'            Ssql = Ssql & vbCrLf & " FECHA )"
'            Ssql = Ssql & vbCrLf & " VALUES (" & Cod_Remito & "," & Cod_Tipo_Remito & "," & Cod_Operacion_Remito & ","
'            Ssql = Ssql & vbCrLf & COD_TIPO_ALMACENAMIENTO & "," & Elementos.Item(i) & "," & Cod_CLiente & ","
'            Ssql = Ssql & vbCrLf & Fecha & ")"
'            ExecutarSql Ssql
'        Next
End Sub
'Public Sub Guardar_Remito()
'
'Dim Sql As String
'Dim R As Integer
'Dim c As Integer
'Dim oradyn As New ADODB.Recordset
'Dim Proximo_Nro_Remito As Long
'
'On Error GoTo OraError
'
'    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
'            Screen.MousePointer = 11
'            strConBasa , 0 ,1.BeginTrans
'            Proximo_Nro_Remito = ProximoRemito
'            ' INSERTAR EN REMITO CUERPO
'            Dim COD_TIPO_ALMACENAMIENTO As Integer
'            Dim NRO_REMITO As Long
'            Dim TIPO As Integer
'            Dim OPERACION As Integer
'            Dim ESTADO As Integer
'            Dim Fecha As String
'            Dim ID_CLIENTE As Integer
'            Dim OBSERVACIONES As String
'            Dim CANTIDAD As Integer
'            Dim AUDIT_USUARIO As String
'            Dim AUDIT_FECHA As String
'
'
'
'
'             Select Case CRequerimientos.Item(1).TIPO
'             Case 1, 3, 8, 9
'                COD_TIPO_ALMACENAMIENTO = 0
'             Case 2, 4
'                COD_TIPO_ALMACENAMIENTO = 1
'             Case 10, 11
'                COD_TIPO_ALMACENAMIENTO = 3
'             End Select
'
'            NRO_REMITO = Proximo_Nro_Remito
'            TIPO = cboTipoRemito.ItemData(cboTipoRemito.ListIndex)
'            OPERACION = cboRemito_Operacion.ItemData(cboRemito_Operacion.ListIndex)
'            ESTADO = cboRemito_Estados.ItemData(cboRemito_Estados.ListIndex)
'            Fecha = " TO_DATE('" & maskFechaRemito.Text & "','DD/MM/YYYY')"
'            ID_CLIENTE = lblIDCliente.Caption
'            If txtObservaciones.Text <> "" Then
'                OBSERVACIONES = "'" & UCase(txtObservaciones.Text) & "'"
'            Else
'                OBSERVACIONES = "NULL"
'            End If
'            CANTIDAD = lblCantidad.Caption
'            AUDIT_USUARIO = "'" & UserName & "'"
'            AUDIT_FECHA = SysDate
'            REMITO_CUERPO_ADD NRO_REMITO, TIPO, OPERACION, ESTADO, Fecha, ID_CLIENTE, _
'            OBSERVACIONES, CANTIDAD, AUDIT_USUARIO, AUDIT_FECHA, COD_TIPO_ALMACENAMIENTO
'
'
'            ' CAMBIO DE ESTADO REQUERIMIENTO
'            Dim IDPERSONAL As Integer
'            Dim I As Integer
'            Dim Bandera As Boolean
'            Bandera = False
'            Dim RS As New ADODB.Recordset
'            Dim Filtro As String
'            Dim FECHARECEPCION As Date
'            Dim IDTIPOREQUERIMIENTO As Integer
'            For I = 0 To lstPersonal.ListCount - 1
'                IDPERSONAL = Mid(lstPersonal.List(I), 1, 2)
'                If lstPersonal.Selected(I) Then
'                    If Bandera = True Then
'                        CambioEstadoRemito IDPERSONAL, False, 4, 6, lbRequerimiento
'                    Else
'                        CambioEstadoRemito IDPERSONAL, True, 4, 6, lbRequerimiento
'                        Bandera = True
'                    End If
'                End If
'            Next
'            Sql = "Update requerimiento set idremito = " & Proximo_Nro_Remito
'            Sql = Sql & vbCrLf & " where idrequerimiento =" & CRequerimientos.Item(1).NumeroRequerimiento
'            ExecutarSql Sql
'
'            'INSERTAR EN REMITO DELTALLE
'
'             With grdCajasLibros
'                For R = 1 To .Rows - 1
'                    For c = 1 To .Cols - 1
'                        If .TextMatrix(R, c) <> "" Then
'                            REMITO_DETALLE_ADD NRO_REMITO, .TextMatrix(R, c), COD_TIPO_ALMACENAMIENTO
'                            GrabarMovHistorico Proximo_Nro_Remito, .TextMatrix(R, c), lblIDCliente.Caption, .TextMatrix(R, c), TIPO, OPERACION, Fecha, COD_TIPO_ALMACENAMIENTO, AUDIT_USUARIO, SysDate
'                        End If
'                    Next
'                Next
'
'
'           'MOVIMIENTO EN TABLA CONTENEDO
'           Select Case cboTipo_Almacenado.Text
'           Case "CAJA"
'                    Select Case cboTipoRemito.ItemData(cboTipoRemito.ListIndex)
'                    Case 1 'CONSULTA
'                        For R = 1 To grdCajasLibros.Rows - 1
'                            For c = 1 To grdCajasLibros.Cols - 1
'                                If grdCajasLibros.TextMatrix(R, c) <> "" Then
'                                    Sql = "UPDATE CONTENEDOR SET "
'                                    Sql = Sql & vbCrLf & " ESTADO = 3 "
'                                    Sql = Sql & ", NRO_REMITO = " & Proximo_Nro_Remito
'                                    Sql = Sql & ", F_MODIFICACION = " & SysDate
'                                    Sql = Sql & vbCrLf & " WHERE "
'                                    Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
'                                    Sql = Sql & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(R, c))
'                                    Sql = Sql & " AND ESTADO = 2 "
'                                    ExecutarSql Sql
'                                End If
'                            Next
'                        Next
'                    Case 2 'CAJAS VACIAS
'                        For R = 1 To grdCajasLibros.Rows - 1
'                            For c = 1 To grdCajasLibros.Cols - 1
'                                If grdCajasLibros.TextMatrix(R, c) <> "" Then
'                                    Sql = "UPDATE CONTENEDOR SET "
'                                    Sql = Sql & vbCrLf & " ESTADO = 5 "
'                                    Sql = Sql & vbCrLf & ", NRO_REMITO = " & Proximo_Nro_Remito
'                                    Sql = Sql & vbCrLf & ", F_MODIFICACION = " & SysDate
'                                    Sql = Sql & vbCrLf & " WHERE "
'                                    Sql = Sql & vbCrLf & " COD_CLIENTE = " & CLng(lblIDCliente.Caption)
'                                    Sql = Sql & vbCrLf & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(R, c))
'                                    Sql = Sql & vbCrLf & " AND ESTADO = 4 "
'                                    ExecutarSql Sql
'                                End If
'                            Next
'                        Next
'                    End Select
'            Case "LIBRO"
'                    For R = 1 To grdCajasLibros.Rows - 1
'                        For c = 1 To grdCajasLibros.Cols - 1
'                            If grdCajasLibros.TextMatrix(R, c) <> "" Then
'                                Sql = "UPDATE LIBROS SET "
'                                Sql = Sql & vbCrLf & " ESTADO = 3 "
'                                Sql = Sql & ", NRO_REMITO = " & Proximo_Nro_Remito
'                                Sql = Sql & ", AUDIT_FECHA = " & SysDate
'                                Sql = Sql & ", AUDIT_USUARIO = '" & UserName
'                                Sql = Sql & vbCrLf & "' WHERE "
'                                Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
'                                Sql = Sql & " AND NRO_LIBRO_INTERNO = " & CLng(grdCajasLibros.TextMatrix(R, c))
'                                Sql = Sql & " AND ESTADO = 2 "
'                                ExecutarSql Sql
'                            End If
'                        Next
'                    Next
'             Case "LEGAJO"
'                  For R = 1 To grdCajasLibros.Rows - 1
'                        For c = 1 To grdCajasLibros.Cols - 1
'                            If grdCajasLibros.TextMatrix(R, c) <> "" Then
'                                    Sql = " Update LEGAJOS"
'                                    Sql = Sql & vbCrLf & " SET COD_ESTADO = 3,"
'                                    Sql = Sql & vbCrLf & "  COD_REMITO = " & Proximo_Nro_Remito
'                                    Sql = Sql & vbCrLf & " , FECHA = " & SysDate
'                                    Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & CInt(lblIDCliente.Caption)
'                                    Sql = Sql & vbCrLf & " And ID_CLIENTE_LEGAJO = " & CLng(grdCajasLibros.TextMatrix(R, c))
'                                    Sql = Sql & vbCrLf & " AND COD_ESTADO = 2 "
'                                    ExecutarSql Sql
'                            End If
'                        Next
'                    Next
'            End Select
'            End With
'            strConBasa , 0 ,1.CommitTrans
'            MsgBox "El remito fue grabado con exito", vbExclamation, "Remito"
'            ImprimirRemito CLng(Proximo_Nro_Remito)
'            Screen.MousePointer = 0
'            On Error GoTo ErrorPrn
'            frmControlEstados.CargarTreenAÑANA
'            Unload Me
'    End If
'Exit Sub
'OraError:
'    Screen.MousePointer = 0
'    strConBasa , 0 ,1.RollbackTrans
'    frmLogOraError.Show MODAL
'    Exit Sub
'
'ErrorPrn:
'    MsgBox ERROR
'    Exit Sub
'
'End Sub


Public Sub REMITO_CUERPO_ADD(NRO_REMITO As Long, TIPO As Integer, OPERACION As Integer, ESTADO As Integer, Fecha As String, ID_CLIENTE As Integer, _
OBSERVACIONES As String, CANTIDAD As Integer, AUDIT_USUARIO As String, AUDIT_FECHA As String, COD_TIPO_ALMACENAMIENTO As Integer)
Dim sql As String
    sql = "INSERT INTO REMITOS_CUERPO (NRO_REMITO , TIPO , OPERACION, ESTADO, FECHA, ID_CLIENTE ,"
    sql = sql & vbCrLf & " OBSERVACIONES , CANTIDAD , AUDIT_USUARIO , AUDIT_FECHA , COD_TIPO_ALMACENAMIENTO )"
    sql = sql & vbCrLf & " VALUES (" & NRO_REMITO & "," & TIPO & "," & OPERACION & "," & ESTADO & "," & Fecha & "," & ID_CLIENTE & ","
    sql = sql & vbCrLf & OBSERVACIONES & "," & CANTIDAD & "," & AUDIT_USUARIO & "," & AUDIT_FECHA & "," & COD_TIPO_ALMACENAMIENTO & ")"
    ExecutarSql sql

End Sub

Public Sub REMITO_DETALLE_ADD(NRO_REMITO As Long, DESDE As Long, TIPO_ALMACENADO As Integer)
    Dim sql As String
    sql = "INSERT INTO REMITOS_DETALLE  (NRO_REMITO, DESDE, HASTA,  NRO_CAJA, TIPO_ALMACENADO ,AUDIT_USUARIO, AUDIT_FECHA)"
    sql = sql & " VALUES (" & NRO_REMITO & "," & DESDE & "," & DESDE & "," & DESDE & "," & TIPO_ALMACENADO & ",'basa'," & SysDate & ")"
    ExecutarSql sql
 End Sub

Public Sub CopiarDatosGrilla(Grilla As DataGrid)
Dim c As Integer
Dim R As Integer
Dim RSDATOS As ADODB.Recordset
Dim dato As String
Dim ColGrilla As Integer
Set RSDATOS = New ADODB.Recordset

Set RSDATOS.DataSource = Grilla.DataSource
 On Error GoTo salir
 

For c = 0 To RSDATOS.Fields.Count - 1
    dato = dato & RSDATOS.Fields(c).Name & vbTab
 Next
dato = dato & vbCrLf
Do While Not RSDATOS.EOF
 For c = 0 To RSDATOS.Fields.Count - 1
    dato = dato & RSDATOS.Fields.Item(c).Value & vbTab
 Next
    RSDATOS.MoveNext
    dato = dato & vbCrLf
Loop
 Clipboard.Clear
 Clipboard.SetText dato
 MsgBox "LOS DATOS FUERON COPIADOS"
salir:
If Err.Number <> 0 Then
    MsgBox "No se encontraron registros"
    Exit Sub
End If
 
End Sub

