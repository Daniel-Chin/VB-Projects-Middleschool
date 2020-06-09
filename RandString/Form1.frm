VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Randomize
Dim s As String
Dim a As Integer
For k = 1 To 12
    a = Int(Rnd() * 10) '10 or 62
    Select Case a
        Case Is <= 9
            a = a + 48
        Case 10 To 35
            a = a + 65 - 10
        Case Else
            a = a - 36 + 97
    End Select
    s = s & Chr(a)
Next k
Tag = InputBox("", "", s)
End
End Sub

