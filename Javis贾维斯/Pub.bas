Attribute VB_Name = "Pub"
Option Explicit

Public Function OutStr(Chang As String, Duan As String) As Integer
If Len(Duan) <> 1 Then MsgBox "debug"
For OutStr = Len(Chang) To 1 Step -1
    If Mid(Chang, OutStr, 1) = Duan Then Exit Function
Next OutStr
OutStr = 0
End Function

Public Sub Un()
Unload AI
Unload Zhu
End
End Sub

Public Function Max(SA, SB)
If SA < SB Then
    Max = SB
Else
    Max = SA
End If
End Function

Public Function Min(SA, SB)
If SA > SB Then
    Min = SB
Else
    Min = SA
End If
End Function

