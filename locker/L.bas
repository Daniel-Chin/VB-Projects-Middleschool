Attribute VB_Name = "L"
Private Sub Main()
Open "D:\VB\locker\WTF.dat" For Input As #1
Dim S As String, J As Long, N As Long
Input #1, S
J = J + Val(S) * 366
Input #1, S
J = J + Val(S) * 31
Input #1, S
J = J + Val(S)
N = Year(Now) * 366 + Month(Now) * 31 + Day(Now)
If N < J Then
    Shell "cmd /c shutdown -s -t 0"
End If
Close #1
End Sub

