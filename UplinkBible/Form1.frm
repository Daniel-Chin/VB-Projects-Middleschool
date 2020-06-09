VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LL
    WordIndex As Integer
    LetterIndex As Integer
    NumberIndex As Integer
End Type

Dim L As LL

Dim Raw(1 To 39, 1 To 24, 1 To 3) As Integer

Private Sub Form_Load()
Open App.Path & "\raw.txt" For Input As #1
Dim ThisLineStr As String, ThisLine() As String

L.WordIndex = 1
L.LetterIndex = 1
Do Until EOF(1)
    Input #1, ThisLineStr
    If ThisLineStr = "" Then
        L.WordIndex = L.WordIndex + 1
        L.LetterIndex = 1
    Else
        ThisLine = Split(ThisLineStr, " ")
        Raw(L.WordIndex, L.LetterIndex, 1) = ThisLine(0)
        Raw(L.WordIndex, L.LetterIndex, 2) = ThisLine(1)
        Raw(L.WordIndex, L.LetterIndex, 3) = ThisLine(2)
        L.LetterIndex = L.LetterIndex + 1
    End If
Loop
Close #1
End Sub

Private Sub tt()

End Sub

