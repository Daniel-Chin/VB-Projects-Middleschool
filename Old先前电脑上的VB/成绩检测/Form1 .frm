VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "成绩检测"
   ClientHeight    =   3984
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10788
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3984
   ScaleWidth      =   10788
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "查看彩蛋列表"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8160
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "指环王"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      TabIndex        =   4
      Top             =   3204
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8160
      TabIndex        =   2
      Text            =   "在此输入秘籍"
      Top             =   3120
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   25.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   1680
      TabIndex        =   0
      Text            =   "0"
      Top             =   3120
      Width           =   6372
   End
   Begin VB.Label Label1 
      Caption         =   "成绩："
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim caidan As Double

Private Sub Command1_Click()
MsgBox "以下是彩蛋列表：我记忆力不差；我要退出；我靠，怎么老是零蛋；你是谁；你很无聊；秦楠枫好帅；尼玛；操；四色定理；指环王；彩蛋。另外：下一次进入本程序时只要在秘籍输入框中输入“SAGA”即可。", vbOKOnly, "你记忆力真差！"
End Sub

Private Sub text1_Change()
Form1.ForeColor = RGB(0, 0, 0)
a = Val(Text1)
Cls
If a = 0 Then Print "零蛋"
If a > 0 And a < 60 Then Print "不及格"
If a >= 60 And a < 90 Then Print "合格"
If a >= 90 And a < 100 Then Print "优秀"
If a = 100 Then Print "满分"
If a > 100 Or a < 0 Then Print "请输入正常的成绩！"
If Text1 = "我要退出" Then MsgBox "是你说要退出的！", , "退出！": End
If Text1 = "我靠，怎么老是零蛋" Then Cls: Call 彩蛋: Print "笨蛋，你自己不输数字！": Call 彩蛋
If Text1 = "彩蛋" Then Cls: Print "你以为世界上所有": Print "程序都有彩蛋么？"
If Text1 = "你很无聊" Then Cls: Call 彩蛋:    Print "你不无聊么？": Call 彩蛋
If Text1 = "数字" Then Cls: Call 彩蛋:    Print "我叫你输数字你就输数字。。": Call 彩蛋
If Text1 = "秦楠枫好帅" Then Cls:  Call 彩蛋:   Print "谢谢，你也很帅！": Call 彩蛋
If Text1 = "尼玛" Then Cls: Call 彩蛋:    Print "不许骂人！": Call 彩蛋
If Text1 = "操" Then Cls: Call 彩蛋:    Print "文明点，说“操令堂！”": Call 彩蛋
If Text1 = "操令堂" Then Cls: Call 彩蛋:    Print "真听话。": Call 彩蛋
If Text1 = "指环王" Then Cls: Call 彩蛋:    Print "Speak Friend and enter！": Call 彩蛋
If Text1 = "Friend" Then Cls: Call 彩蛋:    Print "真聪明！（但是正确答案": Print "往往不止一个）": Call 彩蛋
If Text1 = "Mellon" Then Cls: Call 彩蛋:    Print "看来你已经通读《指环王》了！": Call 彩蛋: Command2.Visible = True
If Text1 = "你是谁" Then Cls:  Call 彩蛋:   Print "我的名字叫秦楠枫。": Call 彩蛋
If Text1 = "终极隐藏" Then Cls: Call 彩蛋:    Print "中华人民共和国万岁！！！": Call 彩蛋
If Text1 = "我记忆力不差" Then Cls: Call 彩蛋:    Print "记忆力不差就别点按钮！": Call 彩蛋
If Text1 = "四色定理" Then Cls: Call 彩蛋:    Print "神奇的邮戳！": Call 彩蛋
If Text1 = "Four is enough" Then Cls: Call 彩蛋:    Print "原来你是拓扑兼历史尖子生！": Call 彩蛋
End Sub

Private Sub 彩蛋()
If Text1 = "终极隐藏" Then Form1.ForeColor = RGB(255, 0, 0) Else Form1.ForeColor = RGB(0, 0, 0)
Command1.Visible = True
If caidan = 1 Then
Else
    caidan = 1
    Call MsgBox("恭喜你发现彩蛋。以下是彩蛋列表：我要退出；我靠，怎么老是零蛋；你是谁；你很无聊；秦楠枫好帅；尼玛；操；四色定理；彩蛋；指环王。另外：下一次进入本程序时只要在秘籍输入框中输入“SAGA”即可。", 0, "彩蛋")
End If
End Sub

Private Sub Text2_Change()
If Text2 = "SAGA" Then Call 彩蛋
End Sub

Private Sub Command2_Click()
    MsgBox "Come in soon!（这是一个终极隐藏彩蛋）", , "我也很喜欢指环王。"
End Sub
