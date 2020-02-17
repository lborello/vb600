VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   900
      TabIndex        =   1
      Top             =   1980
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   675
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim MyName
Dim MyPath
Dim DateMax As Date
MyPath = "\\Recepcion\FAXserve\users\0__Administrador\"
MyName = Dir("\\Recepcion\FAXserve\users\0__Administrador\*.dcx", vbDirectory)
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
      
    If DateMax < CDate(FileDateTime(MyPath & MyName)) Then
        MsgBox TimeValue(CDate(FileDateTime(MyPath & MyName)))
     End If
     DateMax = FileDateTime(MyPath & MyName)
     Debug.Print MyName   ' Display entry only if it
     Debug.Print FileDateTime(MyPath & MyName)
        
      
   End If
   MyName = Dir   ' Get next entry.
Loop
End Sub

Private Sub Command2_Click()
Dim MyName
Dim MyPath
Dim DateMax As Date
Dim MaxTime As Timer
Dim DateReq As Date
Dim Timereq As Date
MousePointer = 11

MyPath = "\\Recepcion\FAXserve\users\0__Administrador\"
MyName = Dir(MyPath & "*.dcx", vbDirectory)
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
    DateReq = format(CDate(FileDateTime(MyPath & MyName)), "DD/MM/YYYY")
    Timereq = TimeValue(FileDateTime(MyPath & MyName))
    Debug.Print "Nombre: " & MyName & "  Dia:" & DateReq; " Hora:" & Timereq
      If DateReq = "17/05/2000" Then
         MsgBox "lll"
      End If
      
   End If
   MyName = Dir   ' Get next entry.
Loop
MousePointer = 0
End Sub


