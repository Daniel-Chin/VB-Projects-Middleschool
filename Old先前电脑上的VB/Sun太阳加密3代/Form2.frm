VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "请输入密码"
   ClientHeight    =   564
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   564
   ScaleWidth      =   3624
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   3612
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.Height = Text1.Height
Form2.Width = Text1.Width
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "sh098-01B" Then
        Form1.Visible = True
        Unload Form2
    Else
        If MsgBox("请输入正确的密码。", vbCritical + vbRetryCancel, "密码错误") = vbRetry Then
            Text1 = ""
            Text1.SetFocus
        Else
            End
        End If
    End If
End If
End Sub
