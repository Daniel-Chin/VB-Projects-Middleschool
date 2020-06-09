VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "设置密码"
   ClientHeight    =   5004
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7644
   LinkTopic       =   "Form3"
   ScaleHeight     =   5004
   ScaleWidth      =   7644
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   1296
      TabIndex        =   2
      Top             =   2880
      Width           =   5052
   End
   Begin VB.TextBox Text1 
      Height          =   612
      Left            =   2496
      TabIndex        =   1
      Top             =   2040
      Width           =   2652
   End
   Begin VB.Label Label1 
      Caption         =   "　　您是第一次使用本程序，请设置密码。密码设定之后不可更改。若遗失密码，请联系程序作者 I_am_qnf@163.com。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   714
      TabIndex        =   0
      Top             =   360
      Width           =   6216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mima As String
Private Sub Command1_Click()
    If InputBox("请再次输入密码。", "密码确认", "") = Mima Then
        Form1.Visible = True
        XieRu Mima
        MsgBox "注册成功！", vbInformation + vbOKOnly, "恭喜"
        End
    Else
        MsgBox "两次密码输入不符。", vbCritical, "错误"
    End If
End Sub
Private Sub Text1_Change()
Mima = Text1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
Private Sub XieRu(PW As String)
Open App.Path & "\。情书加密\delper\1.txt" For Output As #2
    Print #2, PW
Close #2
End Sub
