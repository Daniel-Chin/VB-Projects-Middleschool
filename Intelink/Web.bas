Attribute VB_Name = "Web"
Option Explicit

Const JieGuo As String = "百度为您找到相关结果约"
Dim File As String

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

Private Function Download(ByVal URL As String) As Boolean
   Dim lngReturn As Long
   lngReturn = URLDownloadToFile(0, URL, File, 0, 0)
   If lngReturn = 0 Then Download = True
End Function

Public Function Check(ShuRu As String, GeiDing As String) As Integer
Dim A As Long, B As Long, C As Long
Dim Bili As Double
B = Search(ShuRu & "AND" & GeiDing)
A = Search(ShuRu)
C = Search(GeiDing)
If A = 0 Or C = 0 Then
    Check = -999
Else
    Bili = B / Sqr(A) / Sqr(C)
    If Bili = 0 Then
        Check = -999
    Else
        Check = Log(Bili) + 10
    End If
End If
If Check < 0 Then Check = 0
End Function

Public Sub Initialize()
File = App.Path & "\temp.txt"
End Sub

Public Function Search(Key As String) As Long
Dim Hang As String
If Not Download("http://www.baidu.com/s?wd=" & Key) Then
    GoTo Err
End If
Hang = LuanMa.Convert(File)
    If InStr(Hang, JieGuo) = 0 Then
Err:    MsgBox "搜索引擎失效！", vbCritical + vbOKOnly, "ERROR"
        Unload Form1
        End
    Else
        Hang = Mid(Hang, InStr(Hang, JieGuo), 999)
        Hang = Mid(Hang, Len(JieGuo) + 1, 19)
        Hang = Left(Hang, InStr(Hang, "个") - 1)
        Search = Format(Hang)
    End If
If Search = 100000000 Then
    Search = 0
End If
End Function
