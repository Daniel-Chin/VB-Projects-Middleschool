Attribute VB_Name = "Module1"
Private Sub Main()
Randomize
Dim F2 As String
Dim F1 As String
F1 = Command
If Len(F1) <= 3 Then
    MsgBox "�ñ�exe���ĵ�����������", vbCritical + vbOKOnly
    End
End If
If Left(F1, 1) = Chr(34) Then
    F1 = Mid(F1, 2, Len(F1) - 2)
End If
Open F1 For Binary As #1
F2 = InputBox("To?", , App.Path & "\COPIED")
Dim k As Long, C As Byte, L As Long, tL As Long
L = LOF(1)
tL = 666666
If tL >= L Then tL = L
F1 = InputBox(F1 & "�ܳ���Ϊ" & L & "���Ƽ�����" & tL & "����Ҫ���ƶ��٣�", , tL)
If F1 = "" Then MsgBox "û�ж�ȡ��": End
L = F1
Open F2 For Binary As #2
For k = 1 To L
    Get #1, k, C
    Put #2, k, C
Next k
Close #1
Close #2
MsgBox "Done."
End Sub
