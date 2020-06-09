VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "购买水果方案"
   ClientHeight    =   6420
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11556
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11556
   StartUpPosition =   3  '窗口缺省
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
      Height          =   6252
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始计算！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6132
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   4812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.FontSize = 10
Text1 = ""
For i = 1 To 50
    s = i * 3 + (50 - i)                                                                      '计算花掉的钱
    If s >= 100 And i > 0 And 50 - i > 0 Then                                                 '检测方案是否合格
        Text1 = Text1 & "可以购买香蕉" & i & "个，苹果" & (50 - i) & "个。" & Chr(13) & Chr(10) '输出方案
        j = j + 1                                                                              '累加方案总数
    End If
Next i
Text1 = Text1 & "共有" & j & "种方案。"                                                        '打印方案总数
End Sub
