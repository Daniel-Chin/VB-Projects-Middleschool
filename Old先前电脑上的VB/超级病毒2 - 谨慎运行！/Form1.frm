VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4932
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9408
   LinkTopic       =   "Form1"
   ScaleHeight     =   4932
   ScaleWidth      =   9408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Sub Form_Load()
    With Me
        .Top = 0
        .Left = 0
    End With
    With Screen
        Me.Width = .Width
        Me.Height = .Height
    End With
    ShowCursor False
End Sub
Private Sub Timer1_Timer()
Dim a As String
Open App.Path & "\1\1.txt" For Input As #1
    Input #1, a
Close #1
If a = "1" Then
    Open App.Path & "\1\1.txt" For Output As #2
        Print #2, "2"
    Close #2
    Kill App.Path & "\2.exe"
    FileCopy App.Path & "\1.exe", App.Path & "\2.exe"
    Shell App.Path & "\2.exe", 1
    End
Else
    Open App.Path & "\1\1.txt" For Output As #2
        Print #2, "1"
    Close #2
    Kill App.Path & "\1.exe"
    FileCopy App.Path & "\2.exe", App.Path & "\1.exe"
    Shell App.Path & "\1.exe", 1
    End
End If
End Sub
