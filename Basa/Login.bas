Attribute VB_Name = "Modulo2"

Option Explicit

Rem Global CONBASA  As CONBASA
Rem Global CONBASA As CONBASA
Rem Global OraSqlStmt As OraSqlStmt

Global EstadoConexion As Boolean
Global MdiChildCount As Integer
Global A_Tipo(4) As String
Global Ruta As String

'Connection Information
Global UserName$
Global Password$
Global DatabaseName$
Global Connect$
Global ConexionODBC As String

Rem Global RemitoTipo As Integer
Global RemitoCod As Long
Global RemitoNuevo As Boolean
Global IndexPrn As Integer
Global Primero As Integer
Global Ultimo As Integer
Rem Global Prn As New cfgprn

Global RemitoModificado As Boolean
Global FormTitulo As String
Global FormularioActivo As Integer
Global Reportes(3) As String

' Show parameters
Global Const MARGEN_INFERIOR = 250
Global Const MARGEN_SUPERIOR = 50
Global Const MARGEN_DERECHO = 50
Global Const MARGEN_IZQUIERDO = 150
Global Const MARGEN_INTERNO = 25

Global Const CANCELAR = 0
Global Const ACEPTAR = 1
Global Const MODAL = 1
Global Const MODELESS = 0


Sub CenterForm(F As Form)

  ' Center the specified form within the screen

  F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Function IsDba() As Boolean
  Dim OraRol As New ADODB.Recordset
  IsDba = False
  OraRol.Open "Select * from USER_ROLE_PRIVS WHERE GRANTED_ROLE='DBA'", ConActiva, 0, 1
  If Not OraRol.EOF Then IsDba = True
End Function
Public Function SendMail(sEmailRecipient As String, sEmailSubject As String, sEmailBody As String, Optional sAttachment1 As String, Optional sAttachment2 As String, Optional sAttachment3 As String)

''-----Send an Email Message using Outlook 98-----
'
''Developers Note:  In References, the Microsoft Outlook 98 Object Model must be selected for this to work
''                  Actual file is MSOUTL85.OLB
'
'Dim emailOutlookApp As Outlook.Application
'Dim emailNameSpace As Outlook.Namespace
'Dim emailFolder As Outlook.MAPIFolder
'Dim emailItem As Outlook.MailItem
'Dim EmailRecipient As Recipient
'
''-----Open Outlook in a background process and the Inbox Folder-----
'Set emailOutlookApp = CreateObject("Outlook.Application")
'Set emailNameSpace = emailOutlookApp.GetNamespace("MAPI")
'Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderInbox)
'
''Enable the next line to actually see Outlook Open
''emailFolder.Display
'
''-----Create a new mail message, set the recipient, subject, and body-----
'Set emailItem = emailOutlookApp.CreateItem(olMailItem)
'Set EmailRecipient = emailItem.Recipients.Add(sEmailRecipient)
'emailItem.Subject = sEmailSubject
'emailItem.Body = sEmailBody
'If sAttachment1 <> "" Then
'    emailItem.Attachments.Add sAttachment1
'End If
'If sAttachment2 <> "" Then
'    emailItem.Attachments.Add sAttachment2
'End If
'If sAttachment3 <> "" Then
'    emailItem.Attachments.Add sAttachment3
'End If
'
'
''-----Send the Email-----
'emailItem.Save
'emailItem.Send
'
''-----Close the Outlook Application------
'Rem emailOutlookApp.Quit
'
''----Inform User of Success-----
'Rem  MsgBox "El Correo fue enviado", vbInformation
'
''-----Clear out the memory space held by variables-----
''Usually unnecessary but a good practice
'
'Set emailNameSpace = Nothing
'Set emailFolder = Nothing
'Set emailItem = Nothing
'Set emailOutlookApp = Nothing

End Function

Public Function IndiceExcel(rsIndice As ADODB.Recordset, ExcelIndice As Excel.Worksheet) As Boolean
Dim Base As Integer
Dim Contenido As String
Dim largo  As String
Dim i As Integer
Dim C As Integer
      Base = Len(rsIndice!Indice) / 3
      Dim R As Integer
      R = 4
      With rsIndice
       Do While Not .EOF
'                If Trim(!Tipo_Indice) = "Documento" Then
'                    Contenido = "DOC:" & Format(!ID_CODIGO_DOCUMENTO, "0000") & " - " & UCase(Trim(!DESCRIPCION))
'                 Else
'                    If chkEnvioTodoDocumentos.value = 1 Then
'                        Contenido = "DOC:" & Format(!ID_CODIGO_DOCUMENTO, "0000") & " - " & UCase(Trim(!DESCRIPCION))
'                    Else
'                      Contenido = UCase(Trim(!DESCRIPCION))
'                    End If
'                 End If
'                 If chkEnvioTodoDocumentos.value = True Then
'                       Contenido = "DOC:" & Format(!ID_CODIGO_DOCUMENTO, "0000") & " - " & UCase(Trim(!DESCRIPCION))
'                 End If

 Contenido = "DOC:" & Format(!ID_CODIGO_DOCUMENTO, "0000") & " - " & UCase(Trim(!Descripcion))

                 Dim ce As Excel.Workbook
                 Rem Dim a As Cells



            largo = Len(!Indice) / 3
            For i = Base To largo
                C = i - Base + 1
                If i <> largo Then
                   ExcelIndice.Cells(R, C) = "- - - - - - - - "
                   ' si funciona ExcelIndice.Cells(R, c).Font.Name = "Wingdings 2"
                Else
                   ExcelIndice.Cells(R, C) = Contenido


                   ExcelIndice.Cells(R, 12) = UCase(Trim(rsIndice!Tipo_Indice))
                   If Not IsNull(rsIndice!HABILITAR_FECHA_DESDE) And (rsIndice!HABILITAR_FECHA_DESDE) = True Then
                        If Not IsNull(rsIndice!TITULO_FECHA_DESDE) Then
                            ExcelIndice.Cells(R, 13) = Trim(rsIndice!TITULO_FECHA_DESDE)
                        Else
                            ExcelIndice.Cells(R, 13) = "SI"
                        End If
                   Else
                        ExcelIndice.Cells(R, 13) = ""
                   End If
                   If Not IsNull(rsIndice!HABILITAR_NRO_DESDE) And (rsIndice!HABILITAR_NRO_DESDE) = True Then
                        If Not IsNull(rsIndice!TITULO_NRO_DESDE) Then
                            ExcelIndice.Cells(R, 14) = Trim(rsIndice!TITULO_NRO_DESDE)
                        Else
                            ExcelIndice.Cells(R, 14) = "SI"
                        End If
                  Else
                  ExcelIndice.Cells(R, 14) = ""
                   End If
                   If Not IsNull(rsIndice!HABILITAR_LETRA_DESDE) And (rsIndice!HABILITAR_LETRA_DESDE) = True Then
                        If Not IsNull(rsIndice!TITULO_LETRA_DESDE) Then
                             ExcelIndice.Cells(R, 15) = Trim(rsIndice!TITULO_LETRA_DESDE)
                        Else
                            ExcelIndice.Cells(R, 15) = "SI"
                        End If
                    Else
                     ExcelIndice.Cells(R, 15) = ""
                   End If
                End If
           Next
          R = R + 1

          rsIndice.MoveNext
      Loop
     End With
     IndiceExcel = True
End Function
Public Function ProcesarPorCajas(RsReferenciasCajas As ADODB.Recordset, ExcelReferenciaCaja As Excel.Worksheet) As Boolean
        Dim Titulo As String
        Dim R As Long
        R = 4
        With ExcelReferenciaCaja
            Do While Not RsReferenciasCajas.EOF
            Rem     Debug.Assert RsReferenciasCajas!NRO_CAJA > 6036
                Titulo = Format(RsReferenciasCajas!ID_CODIGO_DOCUMENTO, "0000") & " - " & Trim(RsReferenciasCajas!DESCRIPCIONINDICE)
                ExcelReferenciaCaja.Cells(R, 1) = Titulo
                .Cells(R, 2) = RsReferenciasCajas!NRO_CAJA
                .Cells(R, 3).NumberFormat = "@"
                .Cells(R, 3) = UCase(Trim(RsReferenciasCajas!Descripcion))
                .Cells(R, 4).NumberFormat = "DD/MM/YYYY"
                .Cells(R, 4) = RsReferenciasCajas!FECHA_DESDE
                .Cells(R, 5).NumberFormat = "DD/MM/YYYY"
                .Cells(R, 5) = RsReferenciasCajas!FECHA_HASTA

                If Not IsNull(RsReferenciasCajas!NRO_DESDE) Then
                    .Cells(R, 6) = RsReferenciasCajas!NRO_DESDE
                End If
                If Not IsNull(RsReferenciasCajas!NRO_HASTA) Then
                    .Cells(R, 7) = RsReferenciasCajas!NRO_HASTA
                End If
                .Cells(R, 8) = RsReferenciasCajas!LETRA_DESDE
                .Cells(R, 9) = RsReferenciasCajas!LETRA_HASTA
                Rem .Cells(R, 10) = RsReferenciasCajas!APELLIDO_NOMBRE
                Rem .Cells(R, 11) = RsReferenciasCajas!EXPEDIENTE
                .Cells(R, 12) = RsReferenciasCajas!COD_ID_REFERENCIA
                RsReferenciasCajas.MoveNext
                R = R + 1
            Loop

     End With
End Function
Public Function ProcesarPorIndices(RsReferenciasIndice As ADODB.Recordset, ExcelReferencia As Object) As Boolean
        Dim R As Long
        R = 4
        Dim ID_CODIGO_DOCUMENTO_ANTERIOR As String
       Dim TituloHeredado As String
       With RsReferenciasIndice
            Do While Not RsReferenciasIndice.EOF
                If IsNull(!TituloHerencia) Then
                    TituloHeredado = "DOC:" & Format(!ID_CODIGO_DOCUMENTO, "0000") & "_________________________"
                Else
                    TituloHeredado = "DOC:" & Format(!ID_CODIGO_DOCUMENTO, "0000") & "  //" & Trim(!TituloHerencia)
                End If
                If ID_CODIGO_DOCUMENTO_ANTERIOR = !ID_CODIGO_DOCUMENTO Then



                Else
                    ExcelReferencia.Cells(R, 1) = TituloHeredado
                    R = R + 1
                    ID_CODIGO_DOCUMENTO_ANTERIOR = !ID_CODIGO_DOCUMENTO
                End If

                ExcelReferencia.Cells(R, 1) = TituloHeredado
                ExcelReferencia.Cells(R, 2) = !NRO_CAJA
                ExcelReferencia.Cells(R, 3).WrapText = True
                ExcelReferencia.Cells(R, 3) = UCase(Trim(!Descripcion))
                ExcelReferencia.Cells(R, 4).NumberFormat = "DD/MM/YYYY"
                ExcelReferencia.Cells(R, 4) = !FECHA_DESDE
                ExcelReferencia.Cells(R, 5).NumberFormat = "DD/MM/YYYY"
                ExcelReferencia.Cells(R, 5) = !FECHA_HASTA
                If Not IsNull(!NRO_DESDE) Then
                    ExcelReferencia.Cells(R, 6) = !NRO_DESDE
                End If
                If Not IsNull(!NRO_HASTA) Then
                    ExcelReferencia.Cells(R, 7) = !NRO_HASTA
                End If
                ExcelReferencia.Cells(R, 8) = !LETRA_DESDE
                ExcelReferencia.Cells(R, 9) = !LETRA_HASTA
                ExcelReferencia.Cells(R, 10) = !APELLIDO_NOMBRE
                If Not IsNull(!EXPEDIENTE) Then
                    ExcelReferencia.Cells(R, 11) = !EXPEDIENTE
                Else
                    ExcelReferencia.Cells(R, 11) = ""
                End If

                ExcelReferencia.Cells(R, 12) = !COD_ID_REFERENCIA
                .MoveNext
                R = R + 1
            Loop
      End With


End Function



