Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

Private Sub Main()
Dim yE As Byte, mO As Byte, dA As Byte
'On Error Resume Next
yE = Right(Year(Now), 2)
mO = Month(Now)
dA = Day(Now)
Dim fPath As String, LocalP As String, XiaoBug As Byte, Fail As Boolean
Fail = True
Do While Fail
    XiaoBug = XiaoBug + 1
    LocalP = App.Path & "\bin\" & yE & mO & dA & ".mp3"
    fPath = "http://www.huanqiukexue.com/uploads/sss/60s" & yE & Format(mO, "00") & Format(dA, "00") & ".mp3"
    If dA = 1 Then
        If mO <> 1 Then
            mO = mO - 1
            dA = 31
        Else
            yE = yE - 1
            mO = 12
            dA = 31
        End If
    Else
        dA = dA - 1
    End If
    If XiaoBug = 36 Then
        MsgBox "60sœ¬‘ÿ ß∞‹°£«ÎºÏ≤ÈÕ¯¬Áª∑æ≥°£", vbCritical + vbOKOnly, "¥ÌŒÛ"
        End
    End If
    If Dir(LocalP) <> "" Then End
    Fail = URLDownloadToFile(0, fPath, LocalP, 0, 0)
Loop
Open App.Path & "\bin\1.txt" For Output As #1
Print #1, LocalP
Close #1
End Sub

