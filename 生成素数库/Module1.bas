Attribute VB_Name = "Module1"
Option Explicit

Const Edge As Long = 6666666
Dim k As Long
Dim Z(0 To 999999) As Long
Dim ZTop As Long

Private Sub Main()
Z(0) = 2
Z(1) = 3
Dim GB As Long
GB = 1
Dim kk As Long
Dim IsZ As Boolean
Dim Pct As Integer
For k = 5 To Edge * 2
    IsZ = True
    For kk = 2 To Int(Sqr(k))
        If k Mod kk = 0 Then
            IsZ = False
            Exit For
        End If
    Next kk
    If IsZ Then
        GB = GB + 1
        Z(GB) = k
    End If
    If Int(k / Edge / 2 * 10000) <> Pct Then
        Pct = Int(k / Edge / 2 * 10000)
        Debug.Print "进度：" & Pct / 100 & "%"
    End If
Next k
ZTop = GB
MsgBox "一共找到比" & Edge * 2 & "小的素数" & ZTop & "个，最大的是" & Z(ZTop), vbInformation + vbOKOnly, ZTop
End Sub
