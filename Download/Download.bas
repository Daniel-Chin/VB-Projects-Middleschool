Attribute VB_Name = "Download"
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

Private Sub main()
Dim Url As String, Pth As String
Url = InputBox("��ַ", App.EXEName, "http://")
If Len(Url) <= 3 Then
    End
End If
Pth = InputBox("Ŀ��·������β׺��", App.EXEName, "C:\users\wx\desktop\down.html")
If Len(Pth) <= 3 Then
    End
End If
If URLDownloadToFile(0, Url, Pth, 0, 0) = 0 Then
    If MsgBox("������ɡ���Ŀ���ļ���", vbYesNo + vbInformation, "OK") = vbYes Then
        Shell "explorer " & Pth
    End If
Else
    MsgBox "����ʧ�ܡ���������", vbOKOnly + vbCritical, "Error"
End If
End Sub
