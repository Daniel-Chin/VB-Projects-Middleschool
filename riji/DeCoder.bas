Attribute VB_Name = "Module1"
Option Explicit
Const Temp As String = "D:\VB\riji\Temp"
    Dim SM(1 To 99) As String
    Dim Kit As Integer
    Dim Size As Integer

Private Sub Main()
If InputBox("") = "hvn" Then
    Dim TP As String
    Dim Rt As String
    Rt = "d:\sun\bin\"
    TP = Dir(Rt)
    Do Until TP = ""
        Kit = Kit + 1
        SM(Kit) = Rt & TP
        TP = Dir
    Loop
    Size = Kit
    On Error Resume Next
    Kill App.Path & "\output.txt"
    Open App.Path & "\output.txt" For Output As #9
    For Kit = 1 To Size
        Dim DateM() As String
        DateM() = Split(SM(Kit), ".")
        DateM() = Split(DateM(0), "\")
        DateM() = Split(DateM(UBound(DateM)), "-")
        Print #9, Val(DateM(1)) & "/" & Val(DateM(2))
        Print #9, Reed(SM(Kit))
    Next Kit
    Close #9
    Shell "notepad " & App.Path & "\output.txt", vbMaximizedFocus
End If
End Sub

Private Function Reed(Path As String) As String
Debug.Print Path
Open Path For Binary As #1
Dim S As String
        If 0 Then
ErrH:         MsgBox "Error"
        End If
Kill Temp
Open Temp For Binary As #2
Dim C As Byte
Dim k As Long
For k = 1 To LOF(1) - 1
    Get #1, k, C
    Put #2, k, 255 - C
Next k
Close #1
Close #2
Open Temp For Input As #2
Do Until EOF(2)
    Input #2, S
    Reed = Reed & S & Chr(13) & Chr(10)
Loop
Close #2
End Function

Private Sub listfiles()
Dim k
For k = 1 To Size
    Debug.Print SM(k)
Next k
End Sub
