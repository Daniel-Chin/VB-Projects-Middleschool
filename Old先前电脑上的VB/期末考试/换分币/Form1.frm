VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "换分币"
   ClientHeight    =   6192
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8952
   LinkTopic       =   "Form1"
   ScaleHeight     =   6192
   ScaleWidth      =   8952
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5892
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6252
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始计算！"
      Height          =   5892
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
j = 0
Text1 = ""
For i = 0 To 10
    Text1 = Text1 & "可以换成" & i & "个五角的，和" & 50 - 5 * i & "个一角的。" & Chr(13) & Chr(10)   '输出答案
    j = j + 1                                                                                       '累加器
Next i
Text1 = Text1 & "共有" & j & "种方案。" & Chr(13) & Chr(10)                                          '输出累加玩的方案总数
End Sub
