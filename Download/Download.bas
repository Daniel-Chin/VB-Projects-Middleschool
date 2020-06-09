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
Url = InputBox("网址", App.EXEName, "http://")
If Len(Url) <= 3 Then
    End
End If
Pth = InputBox("目标路径。含尾缀名", App.EXEName, "C:\users\wx\desktop\down.html")
If Len(Pth) <= 3 Then
    End
End If
If URLDownloadToFile(0, Url, Pth, 0, 0) = 0 Then
    If MsgBox("下载完成。打开目标文件吗？", vbYesNo + vbInformation, "OK") = vbYes Then
        Shell "explorer " & Pth
    End If
Else
    MsgBox "下载失败。请检查网速", vbOKOnly + vbCritical, "Error"
End If
End Sub
