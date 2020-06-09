VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Minecraft学习"
   ClientHeight    =   2460
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4752
   LinkTopic       =   "Form2"
   ScaleHeight     =   2460
   ScaleWidth      =   4752
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定嗯"
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4572
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Text            =   "5"
      Top             =   1200
      Width           =   4452
   End
   Begin VB.Label Label2 
      Caption         =   "这是成为MC高手的秘诀！"
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
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   3612
   End
   Begin VB.Label Label1 
      Caption         =   "请设置提醒周期：（分钟）"
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
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4332
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.time = Val(Text1)
Form1.Visible = True
Form2.Visible = False
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub
