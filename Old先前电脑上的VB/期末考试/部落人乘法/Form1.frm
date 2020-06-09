VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "部落人乘法"
   ClientHeight    =   5256
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   ScaleHeight     =   5256
   ScaleWidth      =   9132
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "待运算"
      Top             =   4440
      Width           =   4812
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "待运算"
      Top             =   3240
      Width           =   4812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   4092
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4920
      TabIndex        =   2
      Text            =   "15"
      Top             =   1440
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   1
      Text            =   "13"
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   6
      Top             =   4440
      Width           =   252
   End
   Begin VB.Label Label3 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   252
   End
   Begin VB.Label Label2 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "这是一个乘法器。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim first_ck As Boolean
first_ck = True                '第一次累加
Text3 = ""
i = Val(Text1)
p = Val(Text2)
Do While i >= 1                '除下来>1
    If i Mod 2 = 1 Then        '左列若为奇数
        s = s + p              '右列累加累加
        If first_ck = True Then        '处理“+”号问题
            Text3 = Text3 & p          '生成加法算式
        Else
            Text3 = Text3 & "+" & p
        End If
        first_ck = False
    End If
    i = Int(i / 2)             '左列减半
    p = 2 * p                  '右列加倍
Loop
Text4 = s                      '答案输出
End Sub

Private Sub Form_Load()
Call Command1_Click
End Sub
