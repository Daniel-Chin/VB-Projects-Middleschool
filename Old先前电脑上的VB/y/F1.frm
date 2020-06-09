VERSION 5.00
Begin VB.Form F1 
   BackColor       =   &H80000008&
   ClientHeight    =   5472
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8976
   LinkTopic       =   "Form1"
   ScaleHeight     =   5472
   ScaleWidth      =   8976
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C(1 To 4096, 1 To 4096) As Integer
Private Sub Form_Load()
Dim i As Long, q As Integer, t As Integer, u As Byte, L As Integer, X As Byte
Call asd
End
For L = 1 To 1
    Open "F:\y.bmp" For Binary As #L
    X = (2 ^ (9 - L) - 1)
    For i = 1 To 4096
        For q = 1 To i
            If q = 1 Or q = i Then
                C(i, q) = 1
            Else
                C(i, q) = (C(i - 1, q) + C(i - 1, q - 1)) Mod (2 ^ L)
            End If
            For n = 0 To 2
                Put #L, ((i - 1) * 4096 + q - 1) * 3 + 55 + n, C(i, q) * X
            Next n
        Next q
    Next i
    Close #L
Next L
End
End Sub

Sub asd()
    Open "F:\y.bmp" For Binary As #1
    Dim Code As Byte, i As Long
    For i = 1 To LOF(1)
    Get #1, 55 + i, Code
    Put #1, ((i - 1) * 4096 + q - 1) * 3 + 55 + n, C(i, q) * X

End Sub
