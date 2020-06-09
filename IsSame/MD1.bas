Attribute VB_Name = "MD1"
Option Explicit

Private Sub Main()
MsgBox "this EXE must be run in debug mode. "
Stop
Dim a As String, b As String
a = "D:\Program Files\WatchDogs\Watch_Dogs\bin\3dmGameDll.DLL"
b = "D:\Program Files\WatchDogs\Watch_Dogs\bin\3dmGameDll.txt"
Open a For Binary As #1
Open b For Binary As #2
Dim L As Long
L = LOF(1)
If L = LOF(2) Then
    Dim tImee As Single
    tImee = L / 40000
    MsgBox "俩文件长度为 " & L & "。估计需要时间" & tImee & " 秒。"
    If tImee >= 120 Then
        MsgBox "所需时间挺长，为" & tImee & "秒。", vbExclamation, "警告"
    End If
    Dim k As Long
    Dim c As Byte, d As Byte
    For k = 1 To L
        Get #1, k, c
        Get #2, k, d
        If c <> d Then
            MsgBox "Mismatch!!!"
            Stop
            Close #1, #2
            End
        End If
        If k Mod 166666 = 0 Then
            DoEvents
            Beep
        End If
    Next k
    MsgBox "They are the SAME!!!"
    Stop
Else
    MsgBox "Length not match. They aint the same."
    Stop
End If
Close #1, #2
End Sub
