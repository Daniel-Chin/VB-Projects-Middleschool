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
Private Sub Form_Load()
x = Array(5, 14, 15, 23, 28, 32, 35, 38)
y = Array(20, 38, 2, 22, 41, 3, 16, 40)
'Dim s(0 To 7) As Integer
Dim Dis As Integer
Dim St As String
Dim BestSt As String, Best As Integer
'1 to 7
Best = 9999
For w1 = 1 To 1
For w2 = 1 To 2
For w3 = 1 To 3
For w4 = 1 To 4
For w5 = 1 To 5
For w6 = 1 To 6
For w7 = 1 To 7
St = "1"
St = Left(St, w2 - 1) & "2" & Right(St, 2 - w2)
St = Left(St, w3 - 1) & "3" & Right(St, 3 - w3)
St = Left(St, w4 - 1) & "4" & Right(St, 4 - w4)
St = Left(St, w5 - 1) & "5" & Right(St, 5 - w5)
St = Left(St, w6 - 1) & "6" & Right(St, 6 - w6)
St = Left(St, w7 - 1) & "7" & Right(St, 7 - w7)
St = "0" & St
'Do
Dis = 0
For d = 0 To 7
    prep = Val(Mid(St, d + 1, 1))
    p = Val(Mid(St, d + 2, 1)) Mod 9
    Dis = Dis + Abs(x(p) - x(prep)) + Abs(y(p) - y(prep))
Next d
If Dis < Best Then
    Best = Dis
    BestSt = St
End If
If Dis = 162 Then Debug.Print "aaaa"; St
Next w7
Next w6
Next w5
Next w4
Next w3
Next w2
Next w1
Debug.Print Best; BestSt
End Sub
