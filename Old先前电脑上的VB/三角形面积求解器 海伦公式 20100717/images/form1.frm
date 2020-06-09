VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4464
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "华文新魏"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4464
   ScaleWidth      =   10860
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command11 
      Caption         =   "结束"
      Height          =   612
      Left            =   7560
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command10 
      Caption         =   "继续>>"
      Height          =   372
      Left            =   6120
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton Command9 
      Caption         =   "继续〉〉"
      Height          =   372
      Left            =   6120
      TabIndex        =   19
      Top             =   840
      Width           =   1332
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "显示彩蛋！"
      BeginProperty Font 
         Name            =   "华文琥珀"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3000
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.CommandButton Command7 
      Caption         =   "关于"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   6.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   16
      Top             =   3240
      Width           =   1332
   End
   Begin VB.CommandButton Command5 
      Caption         =   "参加的选修课"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   6.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3000
      TabIndex        =   15
      Top             =   3240
      Width           =   1212
   End
   Begin VB.CommandButton Command6 
      Caption         =   "辅导老师"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   6.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1440
      TabIndex        =   14
      Top             =   3240
      Width           =   1452
   End
   Begin VB.CommandButton Command4 
      Caption         =   "制作人员"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   6.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   732
      Left            =   1440
      TabIndex        =   12
      Top             =   2280
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   732
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始计算"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   1452
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   3612
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   972
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label8 
      Height          =   2172
      Left            =   6120
      TabIndex        =   21
      Top             =   1560
      Width           =   3972
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "form1.frx":0000
      Top             =   3840
      Visible         =   0   'False
      Width           =   10800
   End
   Begin VB.Label Label7 
      Caption         =   "恭喜！你发现了隐藏在本程序中的彩蛋。"
      ForeColor       =   &H00FF0000&
      Height          =   852
      Left            =   6120
      TabIndex        =   18
      Top             =   120
      Width           =   3012
   End
   Begin VB.Label Label6 
      Caption         =   $"form1.frx":12079
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   3000
      TabIndex        =   10
      Top             =   1320
      Width           =   2652
   End
   Begin VB.Label Label5 
      Caption         =   "面积："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   612
   End
   Begin VB.Label Label4 
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   132
   End
   Begin VB.Label Label3 
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   132
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   132
   End
   Begin VB.Label Label1 
      Caption         =   "三角形面积计算器"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   22.2
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = CSng(Text1)
b = CSng(Text2)
c = CSng(Text3)
p = (a + b + c) / 2
If a + b > c Then
    s = Sqr(p * (p - a) * (p - b) * (p - c))
    Text4 = s
Else
    s = "错误！"
End If
End Sub

Private Sub Command10_Click()
Command11.Visible = True
Label8.Caption = "很可惜，本程序的彩蛋部分尚未开发完成。不过，在下一版中彩蛋功能将会到来！谢谢！"
End Sub

Private Sub Command11_Click()
Form1.Width = 6000
Image1.Visible = False
Command8.Visible = False
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
a = 0
b = 0
c = 0
s = 0
Cls
End Sub

Private Sub Command4_Click()
Print "秦楠枫！"
End Sub

Private Sub Command5_Click()
Print "VB程序设计by上外附中！"
End Sub

Private Sub Command6_Click()
Print "陈明老师!"
End Sub

Private Sub Command7_Click()
Command8.Visible = True
End Sub

Private Sub Command8_Click()
Form1.Width = 11076
End Sub

Private Sub Command9_Click()
Image1.Visible = True
Command9.Visible = False
Command10.Visible = True
End Sub

Private Sub Form_Load()
Width = 6000
End Sub
