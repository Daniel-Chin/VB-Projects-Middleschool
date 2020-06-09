VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MC名字修改器"
   ClientHeight    =   4128
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7296
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4128
   ScaleWidth      =   7296
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DL As String

Private Sub Form_Load()
Dim S As String
DL = Chr(13) & Chr(10)
On Error GoTo CuoWu
If 0 Then
CuoWu:      MsgBox "请把本程序与mclauncher.cfg放在同一级菜单下。启动器的文件名最好是“开始游戏.exe”。", vbCritical + vbOKOnly, "启动失败"
            End
End If
Open App.Path & "\mclauncher.cfg" For Input As #1
Dim Ori(1 To 8) As String, yN As String
For i = 1 To 8
    Input #1, Ori(i)
Next i
yN = Right(Ori(2), Len(Ori(2)) - 5)
Dim ShuRu As String
ShuRu = InputBox("本程序专门应对MC启动器设置功能中游戏名字最后一个字符无法消除的问题。输入想要的名字。", "", yN)
Close #1
If ShuRu = "" Then MsgBox "您的名字没有被修改。": End
S = Ori(1) & DL & "游戏名字=" & ShuRu & DL & Ori(3) & DL & Ori(4) & DL & Ori(5) & DL & Ori(6) & DL & Ori(7) & DL & Ori(8) & DL
On Error Resume Next
Kill App.Path & "\mclauncher.cfg"
Open App.Path & "\mclauncher.cfg" For Output As #1
Print #1, S
Close #1
On Error GoTo ZBD
If MsgBox("修改成功。现在运行游戏？", vbYesNo + vbQuestion, "") = vbYes Then
    Shell App.Path & "\开始游戏.exe"
End If
If 0 Then
ZBD:    MsgBox "开始游戏.exe 未找到！单击OK来退出 MC名字修改器。注意您的游戏名已经成功修改了。", vbCritical + vbOKOnly, "错误"
End If
End
End Sub
