VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5592
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5592
   ScaleWidth      =   8580
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
d = 11.1    '
x = -0.1    '
Dim k As Double
Do While True
    s = s + 1
    zx = 999
    jx = (d - x) / 10
    If jx <= 0.0000000000001 Then
        Show
        Me.FontSize = 30
        Print "x=" & d
        Exit Sub
    End If
    For k = x To d Step jx
        y = FC(k)
        y = Abs(y)
        If zx >= y Then
            zx = y
            zxk = k
        End If
    Next k
    x = zxk - jx
    d = zxk + jx
    MsgBox x & "<x<" & d, vbInformation, s
Loop
End Sub

Private Function FC(x As Double) As Double
FC = 11 * x ^ 5 - 404 * x ^ 4 + 4103 * x ^ 3 - 30750 * x ^ 2 + 23760 * x - 101736
End Function

