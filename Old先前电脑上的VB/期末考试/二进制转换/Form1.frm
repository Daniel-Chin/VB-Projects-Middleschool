VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "二进制转换"
   ClientHeight    =   1896
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5328
   LinkTopic       =   "Form1"
   ScaleHeight     =   1896
   ScaleWidth      =   5328
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   1200
      Width           =   2412
   End
   Begin VB.TextBox Text1 
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
      Left            =   2760
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   2412
   End
   Begin VB.Label Label2 
      Caption         =   "输出十进制数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "输入二进制数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
s = 0
a = Val(Text1)                     '取出二位数
Do While s <= Len(CStr(Abs(a)))    '判断循环有没有结束
    b = (a \ 10 ^ s) Mod 10        '截取二位数的指定位数
    If b = 1 Then                  '那个数是1的话
        p = p + 2 ^ (s)            'p累加2的指定位数
    End If
    s = s + 1                      '位数累加
Loop
Text2 = p                          '输出
End Sub
