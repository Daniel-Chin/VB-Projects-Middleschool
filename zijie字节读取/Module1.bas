Attribute VB_Name = "Module1"
Private Sub Main()
Randomize
Dim Nm As String
Dim S As String
S = Command
If Len(S) <= 3 Then
    MsgBox "�ñ�exe���ĵ�����ȡ����", vbCritical + vbOKOnly
    End
End If
If Left(S, 1) = Chr(34) Then
    S = Mid(S, 2, Len(S) - 2)
End If
Open S For Binary As #1
Nm = App.Path & "\" & Rnd() & ".txt"
Dim k As Long, C As Byte, L As Long, tL As Long
L = LOF(1)
tL = 666666
If tL >= L Then tL = L
S = InputBox(S & "�ܳ���Ϊ" & L & "���Ƽ���ȡ" & tL & "����Ҫ��ȡ���٣�", , tL)
If S = "" Then MsgBox "û�ж�ȡ��": End
L = S
Open Nm For Output As #2
For k = 1 To L
    Get #1, k, C
    S = Format(C, "000 ")
    Print #2, S;
Next k
Close #1
Close #2
Shell "notepad " & Nm, vbMaximizedFocus
End Sub


