VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   5748
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7644
   LinkTopic       =   "Form3"
   ScaleHeight     =   5748
   ScaleWidth      =   7644
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3360
      Top             =   960
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Byte
Dim Password As String, Logon As Boolean
Private Sub Form_KeyPress(KeyAscii As Integer)
Password = Password & KeyAscii
If Password = "67810" Then
    Form1.Visible = True
    Logon = True
    Unload Form3
End If
End Sub

Private Sub Form_Load()
Show
Me.FontSize = 15
End Sub


Private Sub Timer2_Timer()
t = t + 1
Select Case t
    Case 1
        Print "抽取SetUp.dll..."
    Case 2
        Print "抽取Log.dll..."
    Case 3
        Print "解码中..."
    Case 4
        Print "ERROR!"
    Case 5
        Print "资料损坏，程序不能打开。将在3秒后自动关闭。"
    Case 6
        Print "2.5秒后程序关闭。"
    Case 7
        Print "2秒后程序关闭。"
    Case 8
        Print "1.5秒后程序关闭。"
    Case 9
        Print "1秒后程序关闭。"
    Case 10
        Print "0.5秒后程序关闭。"
    Case 11
        Print "程序关闭。"
    Case 12
        End
End Select
End Sub
