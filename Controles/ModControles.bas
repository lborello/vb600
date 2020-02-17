Attribute VB_Name = "ModControles"
Option Explicit
Public CONBASA As ADODB.Connection
Enum ABM
    Altas = 0
    Actualizacion = 1
    Bajas = 2
End Enum
Rem  Public Const strConBasa = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BasaSistema;Data Source=Base"
Rem  Public Const strConBasa = "Provider=SQLOLEDB.1;Password=21877471;Persist Security Info=True;User ID=usuario1;Initial Catalog=BasaSistema;Data Source=BASE"
  
  
  Public strConBasa   As String
  

Public Const PasoFax = "\\Server1basa\fax\"




Public Sub Inicio()

On Error GoTo salir
  Dim cad As String, i As Byte, s As Byte, var As Byte
   Open "Z:\Sistemas\Basa\Configuracion.txt" For Input As #1
     While Not EOF(1) 'Recorre archivo hasta que termine
        Input #1, cad
        s = 1 'Controla inicio de cada cadena
        var = 1 'Control el Campo a asignar
        Select Case Trim(Mid(cad, 1, 24))
        Case "strConBasa"
            strConBasa = Replace(Trim(Mid(cad, 25)), ":", ",")
           
        End Select
       Wend
         Close #1
    Exit Sub

salir:
  MsgBox Err.Description
  
End Sub
