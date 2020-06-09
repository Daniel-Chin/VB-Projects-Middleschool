VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Matlab助手"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Const StPath As String = "C:\Program Files\Matlab\R2012b\toolbox\local\startup.m"
Const StPath As String = "F:\Install files\Program files\M\toolbox\local\startup.m"
'Const MatPath As String = "C:\Program Files\Matlab\R2012b\bin\MatlabZhuShou.mat"
Const MatPath As String = "F:\matlab\Ms\MatlabZhuShou.mat"
Dim Emp()

Private Sub Form_Load()
If MsgBox("请确认您没有对Matlab的Startup做过任何自己的更改，因为本程序会对其作出修改。", vbYesNo + vbQuestion, "Matlab助手") = vbNo Then
    MsgBox ("程序已退出。")
    End
Else
    Emp = Array(77, 65, 84, 76, 65, 66, 32, 53, 46, 48, 32, 77, 65, 84, 45, 102, 105, 108, 101, 44, 32, 80, 108, 97, 116, 102, 111, 114, 109, 58, 32, 80, 67, 87, 73, 78, 44, 32, 67, 114, 101, 97, 116, 101, 100, 32, 111, 110, 58, 32, 83, 117, 110, 32, 65, 112, 114, 32, 50, 55, 32, 50, 48, 58, 53, 52, 58, 48, 50, 32, 50, 48, 49, 52, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 73, 77)
    On Error Resume Next
    Kill StPath
    Kill MatPath
    Dim c As Byte
    Open MatPath For Binary As #1
        For k = 1 To 128
            c = Int(Emp(k - 1))
            Put #1, k, c
        Next k
    Close #1
    Open App.Path & "\1.dat" For Binary As #2
    Open StPath For Binary As #3
    For k = 1 To LOF(2)
        Get #2, k, c
        Put #3, k, c
    Next k
    Close #3
    Close #2
    MsgBox "配置完毕。现在请启动Matlab。按OK退出Matlab助手。", vbInformation + vbOKOnly
    End
End If
End Sub
